VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmRefund 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refund"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15165
   ControlBox      =   0   'False
   Icon            =   "frmRefund.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   330
      Left            =   4575
      TabIndex        =   37
      Top             =   9555
      Visible         =   0   'False
      Width           =   3810
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   11265
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":6852
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":752C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":837E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":91D0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":9AAA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":A384
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":AC5E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":B628
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":BF02
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":C21C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":CAF6
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":D3D0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":DCAA
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":DFC4
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":E89E
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":F178
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":FA52
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":1032C
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":10C06
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":114E0
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":11DBA
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":12694
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":12F6E
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":13848
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":14122
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":149FC
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":152D6
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":15BB0
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":1648A
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":16D64
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":1761A
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":17EF4
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":18346
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":18798
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":1AF4A
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":1C1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":1C4E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRefund.frx":22D48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00400040&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15135
      TabIndex        =   34
      Top             =   -45
      Width           =   15195
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmRefund.frx":23A22
         Stretch         =   -1  'True
         Top             =   30
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  refunds"
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
         Left            =   3480
         TabIndex        =   36
         Top             =   225
         Width           =   4575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Tickets Refund"
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
         TabIndex        =   35
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   360
      Left            =   6015
      TabIndex        =   32
      Top             =   10050
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   330
      Left            =   4935
      TabIndex        =   31
      Top             =   10065
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   345
      Left            =   3810
      TabIndex        =   30
      Top             =   10065
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   315
      Left            =   2730
      TabIndex        =   29
      Top             =   10080
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Left            =   1740
      TabIndex        =   28
      Top             =   10005
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmRefund.frx":242EC
      Left            =   4320
      List            =   "frmRefund.frx":242F6
      TabIndex        =   25
      Text            =   "Yes"
      Top             =   6255
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2220
      Left            =   285
      TabIndex        =   24
      Top             =   5820
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   3916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   21
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
      Caption         =   "Routes"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "StatementTicektsID"
         Caption         =   "StatementTicektsID"
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
         DataField       =   "StatementDetails"
         Caption         =   "StatementDetails"
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
         DataField       =   "DepartDateTime"
         Caption         =   "DepartDateTime"
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
         DataField       =   "Route"
         Caption         =   "Route"
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
         DataField       =   "Company Commission"
         Caption         =   "Company Commission"
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
         DataField       =   "Insurance"
         Caption         =   "Insurance"
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
         DataField       =   "ASF"
         Caption         =   "ASF"
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
         DataField       =   "Terminal Fee"
         Caption         =   "Terminal Fee"
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
         DataField       =   "Meals"
         Caption         =   "Meals"
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
         DataField       =   "TicketAmount"
         Caption         =   "Ticket Amount"
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
         DataField       =   "ReFund"
         Caption         =   "Refund ?"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Yes"
            FalseValue      =   "No"
            NullValue       =   "No"
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Posted"
         Caption         =   "Posted"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Yes"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
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
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Height          =   8880
      Left            =   15
      TabIndex        =   2
      Top             =   570
      Width           =   15195
      Begin VB.TextBox txtPayto 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1545
         TabIndex        =   7
         Top             =   150
         Width           =   3375
      End
      Begin VB.TextBox txtRefundDate 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12315
         TabIndex        =   6
         Top             =   150
         Width           =   2460
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6345
         TabIndex        =   5
         Top             =   165
         Width           =   4065
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   8220
         Left            =   90
         ScaleHeight     =   8190
         ScaleWidth      =   15000
         TabIndex        =   3
         Top             =   585
         Width           =   15030
         Begin VB.TextBox txtNetRefundableAmount1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13275
            Locked          =   -1  'True
            TabIndex        =   101
            Text            =   "0.00"
            Top             =   7725
            Width           =   1590
         End
         Begin VB.Frame Frame4 
            Caption         =   "Refund to Airline / Shipping"
            Height          =   7470
            Left            =   10395
            TabIndex        =   70
            Top             =   90
            Width           =   4560
            Begin VB.TextBox txtEvatAddComm 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   105
               Text            =   "0.00"
               Top             =   6750
               Width           =   1590
            End
            Begin VB.TextBox txtCommEvatPercent 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   102
               Text            =   "0.00"
               Top             =   6360
               Width           =   1590
            End
            Begin VB.TextBox txtRecallComm1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   84
               Text            =   "0.00"
               Top             =   5970
               Width           =   1590
            End
            Begin VB.TextBox txtRefundAmountPax1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   83
               Text            =   "0.00"
               Top             =   5580
               Width           =   1590
            End
            Begin VB.TextBox txtNoShowSurcharge1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   82
               Text            =   "0.00"
               Top             =   4215
               Width           =   1590
            End
            Begin VB.TextBox txtRefundFee1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   81
               Text            =   "0.00"
               Top             =   3825
               Width           =   1590
            End
            Begin VB.TextBox txtInsurance1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   80
               Text            =   "0.00"
               Top             =   2670
               Width           =   1590
            End
            Begin VB.TextBox txtTerminalFee1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   79
               Text            =   "0.00"
               Top             =   2280
               Width           =   1590
            End
            Begin VB.TextBox txtASF1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   78
               Text            =   "0.00"
               Top             =   1890
               Width           =   1590
            End
            Begin VB.TextBox txtTax1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   77
               Text            =   "0.00"
               Top             =   1500
               Width           =   1590
            End
            Begin VB.TextBox txtRefundableAmount1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   76
               Text            =   "0.00"
               Top             =   1110
               Width           =   1590
            End
            Begin VB.TextBox txtUsedPortions1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   75
               Text            =   "0.00"
               Top             =   720
               Width           =   1590
            End
            Begin VB.TextBox txtFare1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   74
               Text            =   "0.00"
               Top             =   330
               Width           =   1590
            End
            Begin VB.TextBox txtEvat1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   73
               Text            =   "0.00"
               Top             =   3060
               Width           =   1590
            End
            Begin VB.TextBox txtServiceFee1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   72
               Text            =   "0.00"
               Top             =   3435
               Width           =   1590
            End
            Begin VB.TextBox txtVoidFee1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   71
               Text            =   "0.00"
               Top             =   4605
               Width           =   1590
            End
            Begin VB.Label Label46 
               BackStyle       =   0  'Transparent
               Caption         =   "EVAT + COMMISSION :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   150
               TabIndex        =   104
               Top             =   6765
               Width           =   2415
            End
            Begin VB.Label Label45 
               BackStyle       =   0  'Transparent
               Caption         =   "% EVAT ON COMMISSION :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   150
               TabIndex        =   103
               Top             =   6435
               Width           =   2685
            End
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "FARE PAID :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   105
               TabIndex        =   99
               Top             =   300
               Width           =   2175
            End
            Begin VB.Label Label42 
               BackStyle       =   0  'Transparent
               Caption         =   "LESS: USED PORTION(S) :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   210
               TabIndex        =   98
               Top             =   750
               Width           =   2475
            End
            Begin VB.Label Label41 
               BackStyle       =   0  'Transparent
               Caption         =   "REFUNDABLE AMOUNT :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   135
               TabIndex        =   97
               Top             =   1155
               Width           =   2895
            End
            Begin VB.Label Label40 
               BackStyle       =   0  'Transparent
               Caption         =   "ADD :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   630
               TabIndex        =   96
               Top             =   1500
               Width           =   855
            End
            Begin VB.Label Label39 
               BackStyle       =   0  'Transparent
               Caption         =   "TAX :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1200
               TabIndex        =   95
               Top             =   1515
               Width           =   1125
            End
            Begin VB.Label Label38 
               BackStyle       =   0  'Transparent
               Caption         =   "ASF :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1215
               TabIndex        =   94
               Top             =   1905
               Width           =   1125
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "TERMINAL FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1185
               TabIndex        =   93
               Top             =   2325
               Width           =   1800
            End
            Begin VB.Label Label36 
               BackStyle       =   0  'Transparent
               Caption         =   "INSURANCE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1230
               TabIndex        =   92
               Top             =   2745
               Width           =   1605
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               Caption         =   "REFUND  FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   465
               TabIndex        =   91
               Top             =   3825
               Width           =   3540
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "NO SHOW SURCHARGE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   465
               TabIndex        =   90
               Top             =   4215
               Width           =   2745
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "NET REFUND AMOUNT  :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   135
               TabIndex        =   89
               Top             =   5535
               Width           =   2685
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "LESS :  RECALL COMMISSION :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   150
               TabIndex        =   88
               Top             =   5925
               Width           =   2805
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "EVAT :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1230
               TabIndex        =   87
               Top             =   3105
               Width           =   1875
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "LESS :  SERVICE  FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   180
               TabIndex        =   86
               Top             =   3435
               Width           =   3540
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "VOID FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   465
               TabIndex        =   85
               Top             =   4620
               Width           =   2115
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Refund to Pax"
            Height          =   7470
            Left            =   5790
            TabIndex        =   40
            Top             =   90
            Width           =   4590
            Begin VB.TextBox txtCommPercent 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   111
               Text            =   "0.00"
               Top             =   6390
               Width           =   1605
            End
            Begin VB.CommandButton cmdSet 
               Caption         =   "Set"
               Height          =   360
               Left            =   1950
               TabIndex        =   110
               Top             =   3825
               Width           =   885
            End
            Begin VB.TextBox txtRecallComm 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   54
               Text            =   "0.00"
               Top             =   5955
               Width           =   1590
            End
            Begin VB.TextBox txtRefundAmountPax 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   53
               Text            =   "0.00"
               Top             =   5565
               Width           =   1590
            End
            Begin VB.TextBox txtNoShowSurcharge 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               TabIndex        =   52
               Text            =   "0.00"
               Top             =   4200
               Width           =   1590
            End
            Begin VB.TextBox txtRefundFee 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               TabIndex        =   51
               Text            =   "0.00"
               Top             =   3810
               Width           =   1590
            End
            Begin VB.TextBox txtInsurance 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   50
               Text            =   "0.00"
               Top             =   2655
               Width           =   1590
            End
            Begin VB.TextBox txtTerminalFee 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   49
               Text            =   "0.00"
               Top             =   2265
               Width           =   1590
            End
            Begin VB.TextBox txtASF 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   48
               Text            =   "0.00"
               Top             =   1875
               Width           =   1590
            End
            Begin VB.TextBox txtTax 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               TabIndex        =   47
               Text            =   "0.00"
               Top             =   1485
               Width           =   1590
            End
            Begin VB.TextBox txtRefundableAmount 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   46
               Text            =   "0.00"
               Top             =   1095
               Width           =   1590
            End
            Begin VB.TextBox txtUsedPortions 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               TabIndex        =   45
               Text            =   "0.00"
               Top             =   705
               Width           =   1590
            End
            Begin VB.TextBox txtFare 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   44
               Text            =   "0.00"
               Top             =   315
               Width           =   1590
            End
            Begin VB.TextBox txtEvat 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   43
               Text            =   "0.00"
               Top             =   3045
               Width           =   1590
            End
            Begin VB.TextBox txtServiceFee 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2850
               TabIndex        =   42
               Text            =   "0.00"
               Top             =   3420
               Width           =   1590
            End
            Begin VB.TextBox txtVoidFee 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2865
               TabIndex        =   41
               Text            =   "0.00"
               Top             =   4575
               Width           =   1590
            End
            Begin VB.Label Label49 
               BackStyle       =   0  'Transparent
               Caption         =   "COMMISSION PERCENT %:"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   120
               TabIndex        =   112
               Top             =   6390
               Width           =   2805
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "FARE PAID :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   75
               TabIndex        =   69
               Top             =   285
               Width           =   2175
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "LESS: USED PORTION(S) :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   180
               TabIndex        =   68
               Top             =   735
               Width           =   2475
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "REFUNDABLE AMOUNT :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   105
               TabIndex        =   67
               Top             =   1140
               Width           =   2895
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "ADD :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   600
               TabIndex        =   66
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "TAX :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1170
               TabIndex        =   65
               Top             =   1500
               Width           =   1125
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "ASF :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1185
               TabIndex        =   64
               Top             =   1890
               Width           =   1125
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "TERMINAL FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1155
               TabIndex        =   63
               Top             =   2310
               Width           =   1800
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "INSURANCE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1200
               TabIndex        =   62
               Top             =   2730
               Width           =   1605
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "REFUND  FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   435
               TabIndex        =   61
               Top             =   3810
               Width           =   3540
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "NO SHOW SURCHARGE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   435
               TabIndex        =   60
               Top             =   4200
               Width           =   2745
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "NET REFUND AMOUNT TO PAX :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   105
               TabIndex        =   59
               Top             =   5520
               Width           =   2685
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "LESS :  RECALL COMMISSION :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   120
               TabIndex        =   58
               Top             =   5910
               Width           =   2805
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "EVAT :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Left            =   1200
               TabIndex        =   57
               Top             =   3090
               Width           =   1875
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "LESS :  SERVICE  FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   150
               TabIndex        =   56
               Top             =   3420
               Width           =   3540
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "VOID FEE :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   435
               TabIndex        =   55
               Top             =   4605
               Width           =   2115
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   8115
            Left            =   0
            ScaleHeight     =   8055
            ScaleWidth      =   5655
            TabIndex        =   12
            Top             =   0
            Width           =   5715
            Begin VB.TextBox txtNumSector 
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
               Left            =   2760
               TabIndex        =   38
               Top             =   7500
               Width           =   2745
            End
            Begin VB.Frame Frame2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "TICKET INFORMATION / DETAILS"
               Enabled         =   0   'False
               ForeColor       =   &H00C00000&
               Height          =   5775
               Left            =   0
               TabIndex        =   16
               Top             =   1635
               Width           =   5670
               Begin VB.TextBox txtTicketType 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2715
                  TabIndex        =   109
                  Top             =   1080
                  Width           =   2790
               End
               Begin VB.TextBox txtPaxName 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2715
                  TabIndex        =   107
                  Top             =   2580
                  Width           =   2790
               End
               Begin VB.TextBox txtNetFare 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2715
                  TabIndex        =   26
                  Text            =   "0.00"
                  Top             =   2205
                  Width           =   2790
               End
               Begin VB.TextBox txtGrossFare 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2715
                  TabIndex        =   22
                  Text            =   "0.00"
                  Top             =   1800
                  Width           =   2790
               End
               Begin VB.TextBox txtAirline 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2715
                  TabIndex        =   20
                  Top             =   675
                  Width           =   2790
               End
               Begin VB.TextBox txtTicketIssuedDate 
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2715
                  TabIndex        =   17
                  Top             =   285
                  Width           =   2790
               End
               Begin VB.Label Label48 
                  BackStyle       =   0  'Transparent
                  Caption         =   "TICKET TYPE :"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   105
                  TabIndex        =   108
                  Top             =   1170
                  Width           =   2625
               End
               Begin VB.Label Label47 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PAX NAME :"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   105
                  TabIndex        =   106
                  Top             =   2595
                  Width           =   2220
               End
               Begin VB.Label Label23 
                  BackStyle       =   0  'Transparent
                  Caption         =   "NET FARE FOR THIS TICKET"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   105
                  TabIndex        =   27
                  Top             =   2145
                  Width           =   2760
               End
               Begin VB.Label Label22 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GROSS FARE FOR THIS TICKET"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   105
                  TabIndex        =   21
                  Top             =   1740
                  Width           =   2760
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "AIRLINE / SHIPPING LINE :"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   105
                  TabIndex        =   19
                  Top             =   645
                  Width           =   2625
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "DATE TICKET / MCO ISSUANCE :"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   105
                  TabIndex        =   18
                  Top             =   240
                  Width           =   2625
               End
            End
            Begin VB.TextBox txtTicketNo 
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
               Left            =   1320
               TabIndex        =   0
               Top             =   945
               Width           =   2745
            End
            Begin LVbuttons.LaVolpeButton cmdFind 
               Height          =   480
               Left            =   4170
               TabIndex        =   1
               Top             =   930
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
               MICON           =   "frmRefund.frx":24303
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
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "NO. OF SECTOR"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   39
               Top             =   7530
               Width           =   2415
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               Caption         =   "SEARCH TICKET #"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   0
               TabIndex        =   14
               Top             =   -30
               Width           =   5670
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "TICKET / MCO NO. (S) :"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   255
               TabIndex        =   13
               Top             =   555
               Width           =   3630
            End
         End
         Begin VB.TextBox txtNetRefundableAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   7740
            Width           =   1590
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "NET REFUNDABLE AMOUNT :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   10395
            TabIndex        =   100
            Top             =   7725
            Width           =   2700
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "NET REFUNDABLE AMOUNT :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5850
            TabIndex        =   11
            Top             =   7740
            Width           =   4230
         End
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
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REFUND DATE :"
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
         Left            =   10515
         TabIndex        =   9
         Top             =   105
         Width           =   2640
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
         Left            =   4935
         TabIndex        =   8
         Top             =   165
         Width           =   1245
      End
   End
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   13455
      TabIndex        =   15
      Top             =   9495
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
      MICON           =   "frmRefund.frx":2431F
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
      Left            =   60
      TabIndex        =   23
      Top             =   9480
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
      MICON           =   "frmRefund.frx":2433B
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
   Begin LVbuttons.LaVolpeButton cmdCancel 
      Height          =   480
      Left            =   12075
      TabIndex        =   33
      Top             =   9495
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "frmRefund.frx":24357
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
End
Attribute VB_Name = "frmRefund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Test Ticket :106417310


Dim RsRefund As ADODB.Recordset
Dim RsStatement As ADODB.Recordset
Dim RsFindTicket As ADODB.Recordset
Dim RsRoutes As ADODB.Recordset
Dim AgentCommission As Double
Dim StatementDetail

Dim myTmpEvat
Dim myTmpServiceFee
Dim myTmpRefundFee
Dim myTmpVoidFee
Dim myTmpNoShowFee

Dim myTmpFare

Dim myPublic_AirShipID
Dim myPublic_AccountID


Dim SQL As String

Function CheckIfNotPaid(ByVal param As Long) As Boolean
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Statement WHERE [TransID]=" & param & " AND ([Down]=False and [Paid]=false) "
With Rst
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
            CheckIfNotPaid = True
          Else
            CheckIfNotPaid = False
    End If
Set Rst = Nothing
End With
End Function

Function SA_Return(ByVal usrTicket As String) As String
On Error GoTo FailSafe_Error
Dim Rs              As New ADODB.Recordset
Dim tmpTransID      As Long

SQL = "SELECT * FROM tbl_StatementDetail WHERE [Ticket No]='" & usrTicket & "'"
cn.BeginTrans
        With Rs
                .Open SQL, cn, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                
                        myPublic_AirShipID = .Fields("Airline").Value
                
                        tmpTransID = .Fields("TransID").Value
                        
                        '//now find the SA no# using tmpTransID
                        SQL = "SELECT * FROM tbl_Statement WHERE [TransID]=" & tmpTransID
                        Rs.Close
                        Set Rs = New ADODB.Recordset
                        With Rs
                                .Open SQL, cn, adOpenKeyset, adLockOptimistic
                                If .RecordCount > 0 Then
                                    SA_Return = .Fields("sNumber").Value
                                    
                                    myPublic_AccountID = .Fields("AccountID").Value
                                End If
                        End With
                End If
        End With
        
cn.CommitTrans
Exit Function
FailSafe_Error:
cn.RollbackTrans
End Function
Sub SA_Update(ByVal strStatement As String, ByVal usrTicket As String, ByVal CurPay As Double, ByVal usrTktAmt As Double, ByVal usrDate As Date, ByVal usrVoucherID As Long)
'On Error GoTo FailSafe_Error

Dim Rs              As New ADODB.Recordset
Dim tmpBal          As Double

cn.BeginTrans
SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & strStatement & "'"
Set RsStatement = New ADODB.Recordset
With RsStatement
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                            tmpBal = CDbl(.Fields("Balance").Value)
            
                                 .Fields("Down").Value = True
                            If (tmpBal - CurPay) <= 0 Then
                                 .Fields("Paid").Value = True
                                 .Fields("Down").Value = False
                            Else
                                 .Fields("Paid").Value = False
                            End If
                                 .Fields("Credit Card Activated").Value = True
                                 .Fields("Balance").Value = (tmpBal - CurPay)
                                 .Update
            End If
End With

SQL = "SELECT * FROM tbl_Refund"
Set Rs = New ADODB.Recordset
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("Refund Date").Value = Format(usrDate, "mm/dd/yyyy")
        .Fields("SA no").Value = strStatement
        .Fields("Ticket No").Value = usrTicket
        .Fields("Ticket Amount").Value = CDbl(usrTktAmt)
        .Fields("Refund Amount").Value = CDbl(CurPay)
        .Fields("AirlineID").Value = myPublic_AirShipID
        .Fields("AccountID").Value = myPublic_AccountID
        .Fields("Pax Name").Value = Me.txtPaxName
        .Fields("VoucherID").Value = usrVoucherID
        .Fields("Ticket Type").Value = Me.txtTicketType
        .Update
End With
cn.CommitTrans
FailSafe_Error:
'cn.RollbackTrans
End Sub

Sub LoadValues(param)
Dim ask As Integer
Set RsRefund = New ADODB.Recordset
        SQL = "SELECT * FROM tbl_StatementDetail WHERE [Ticket No]='" & param & "'"
With RsRefund
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            'Check if already refunded
             If .Fields("Refund").Value Then
                 MsgBox "This Ticket was already refunded!", vbCritical
                 With Me.txtTicketNo
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SetFocus
                        Exit Sub
                 End With
             End If
             
            If CheckIfNotPaid(.Fields(1).Value) Then
            ask = MsgBox("This ticket was not paid! continue will subtract the refunded amount to A/R", vbInformation + vbYesNo)
                    If ask = vbNo Then
                    Me.Option1.Value = False
                         With Me.txtTicketNo
                                .SelStart = 0
                                .SelLength = Len(.Text)
                                .SetFocus
                                Exit Sub
                         End With
                    End If
            'Me.Option1.Value = True
            Else
            Me.Option1.Value = False
            End If
             
           
            Set RsRoutes = New ADODB.Recordset
            
            myTmpEvat = IIf(IsNull(.Fields("EVAT").Value), 0, .Fields("EVAT").Value)
            myTmpServiceFee = IIf(IsNull(.Fields("Service Fee").Value), 0, .Fields("Service Fee").Value)
            myTmpRefundFee = IIf(IsNull(.Fields("Refund Fee").Value), 0, .Fields("Refund Fee").Value)
            myTmpVoidFee = IIf(IsNull(.Fields("Void Fee").Value), 0, .Fields("Void Fee").Value)
            myTmpNoShowFee = IIf(IsNull(.Fields("Noshow Fee").Value), 0, .Fields("Noshow Fee").Value)
            Me.txtCommPercent = IIf(IsNull(.Fields("Commision").Value), 0, .Fields("Commision").Value)
            
            'myTmpFare = IIf(IsNull(.Fields("Fare").Value), 0, .Fields("Fare").Value)
            
            SQL = "SELECT * FROM tbl_StatementTickets WHERE [StatementDetails]=" & .Fields("StatementDetails").Value
            With RsRoutes
                        .Open SQL, cn, adOpenKeyset, adLockOptimistic
                        If .RecordCount > 0 Then
                         Set Me.DataGrid1.DataSource = RsRoutes
                         Call MoveCombo
                        End If
            End With
            
             Me.txtPaxName = .Fields("Name").Value
             Me.txtTicketType = .Fields("Ticket Type").Value
             Me.txtGrossFare = Format(.Fields("Gross").Value, "###,##0.00")
             Me.txtNetFare = Format(.Fields("Net").Value, "###,##0.00")
             Me.txtPayto = .Fields("Name").Value
             
             
             AgentCommission = .Fields("Commision").Value
            Call FindStatement(.Fields(1).Value)
            StatementDetail = .Fields(0).Value
            
            
        Else
            MsgBox "No such Ticket!", vbCritical
            .Close
          Set RsRefund = Nothing
          Exit Sub
        End If
End With
ReCalc
End Sub

Sub UpdateTicket(ByVal param As Long, ByVal Criteria As Boolean)
cn.BeginTrans
SQL = "UPDATE tbl_StatementTickets SET [Refund] =" & Criteria & " WHERE [StatementTicektsID]=" & param
cn.Execute SQL
cn.CommitTrans
End Sub


Sub UpdateStatement(ByVal param As Long, ByVal Criteria As Boolean)
If CheckIfCompleteted(param, False) = True Then
cn.BeginTrans
    SQL = "UPDATE tbl_StatementDetail SET [Refund] =" & Criteria & " WHERE [StatementDetails]=" & param
    cn.Execute SQL
cn.CommitTrans
End If
If Me.Option1.Value = False Then
    Call AppendToVoucher
End If
End Sub

Sub AppendToVoucher()
Dim Rs As New ADODB.Recordset
Dim RsDetail As New ADODB.Recordset

SQL = "SELECT * FROM tbl_Voucher"
With Rs
     .Open SQL, cn, adOpenKeyset, adLockOptimistic
     .AddNew
     .Fields(1).Value = Me.txtPayto
     .Fields(2).Value = Me.txtAddress
     .Fields(3).Value = Format(Now, "mm/dd/yyyy")
     .Fields(6).Value = Format(CDbl(Me.txtNetRefundableAmount), "###,##0.00")
     .Fields("For Refund").Value = True
     .Update
End With

SQL = "SELECT * FROM tbl_VoucherDetails WHERE [VoucherID]=" & Rs.Fields(0).Value
Me.Tag = Rs.Fields(0).Value
With RsDetail
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields(1).Value = Rs.Fields(0).Value
        .Fields(2).Value = Me.txtTicketNo
        .Fields(3).Value = Format(CDbl(Me.txtNetRefundableAmount), "###,##0.00")
        .Update
End With
End Sub

Function CheckIfCompleteted(ByVal param As Long, ByVal Criteria As Boolean) As Boolean
Dim Rstmp As ADODB.Recordset
Set Rstmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_StatementTickets WHERE [StatementDetails]=" & param & " AND [Refund]=" & Criteria
With Rstmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
            
                CheckIfCompleteted = False
            Else
                CheckIfCompleteted = True
            End If
           .Close
        Set Rstmp = Nothing
End With
End Function


Function CheckIfPosted(ByVal param As Long) As Boolean
Dim Rstmp As ADODB.Recordset
Dim tmp As Long

Set Rstmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_StatementTickets WHERE [StatementDetails]=" & param
With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            tmp = .RecordCount
        End If
        .Close
      Set Rstmp = Nothing
End With

Set Rstmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_StatementTickets WHERE [StatementDetails]=" & param & " AND [Refund]=TRUE"
With Rstmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
               If tmp = .RecordCount Then
                    CheckIfPosted = False
                    Else
                    CheckIfPosted = True
              End If
           End If
           .Close
        Set Rstmp = Nothing
End With
End Function




Sub FindStatement(param)
Set RsStatement = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Statement WHERE [TransID]=" & param
With RsStatement
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                Me.txtTicketIssuedDate = .Fields("Date").Value
                Me.txtAirline = FindAirlineName(.Fields("Airline").Value)
               
        End If
        .Close
       Set RsStatement = Nothing
End With

End Sub

Sub FindTicketDetails(ByVal param As Long, ByVal Opt As String)
Dim Fare As Double
Dim Insurance As Double
Dim ASF As Double
Dim TFee As Double
Dim Commission As Double
Dim myEvat As Double

Set RsFindTicket = New ADODB.Recordset
SQL = "SELECT * FROM tbl_StatementTickets WHERE [StatementTicektsID]=" & param

With RsFindTicket
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              Fare = CDbl(.Fields("TicketAmount").Value)
              Insurance = CDbl(.Fields("Insurance").Value)
              ASF = CDbl(.Fields("ASF").Value)
              TFee = CDbl(.Fields("Terminal Fee").Value)
              Commission = (CDbl(AgentCommission) / 100) * Fare
              
            
              If Opt = "Add" Then
                
                Me.txtFare = Format(CDbl(Me.txtFare) + Fare, "###,##0.00")
                
                'Me.txtFare = Format(CDbl(Me.txtFare) + myTmpFare, "###,##0.00")
                
                Me.txtInsurance = Format(CDbl(Me.txtInsurance) + Insurance, "###,##0.00")
                Me.txtASF = Format(CDbl(Me.txtASF) + ASF, "###,##0.00")
                Me.txtTerminalFee = Format(CDbl(Me.txtTerminalFee) + TFee, "###,##0.00")
                Me.txtRecallComm = Format(CDbl(Me.txtRecallComm) + Commission, "###,##0.00")
               
              Else
                Me.txtFare = Format(CDbl(Me.txtFare) - IIf(CDbl(Me.txtFare) >= Fare, Fare, 0), "###,##0.00")
                
                'Me.txtFare = Format(CDbl(Me.txtFare) - IIf(CDbl(Me.txtFare) >= myTmpFare, myTmpFare, 0), "###,##0.00")
                
                Me.txtInsurance = Format(CDbl(Me.txtInsurance) - IIf(CDbl(Me.txtInsurance) >= Insurance, Insurance, 0), "###,##0.00")
                Me.txtASF = Format(CDbl(Me.txtASF) - IIf(CDbl(Me.txtASF) >= ASF, ASF, 0), "###,##0.00")
                Me.txtTerminalFee = Format(CDbl(Me.txtTerminalFee) - IIf(CDbl(Me.txtTerminalFee) >= TFee, TFee, 0), "###,##0.00")
                Me.txtRecallComm = Format(CDbl(Me.txtRecallComm) - IIf(CDbl(Me.txtRecallComm) >= Commission, Commission, 0), "###,##0.00")
            
              End If
              
'simulate a false evat
'just to compute natak an nako

            If myTmpEvat = "" Or myTmpEvat = 0 Then
                Dim tmpVar
                tmpVar = CDbl(Me.txtFare) + CDbl(Me.txtTax) + CDbl(Me.txtASF) + CDbl(txtTerminalFee) + CDbl(txtInsurance)
                myTmpEvat = CDbl(Me.txtNetFare) - tmpVar
            End If
            
            
                Me.txtEvat = Format(myTmpEvat, "###,##0.00")
                Me.txtServiceFee = Format(myTmpServiceFee, "###,##0.00")
                Me.txtRefundFee = Format(myTmpRefundFee, "###,##0.00")
                Me.txtVoidFee = Format(myTmpVoidFee, "###,##0.00")
                Me.txtNoShowSurcharge = Format(myTmpNoShowFee, "###,##0.00")
                
                
'Copy values from pax to airline
Me.txtFare1 = Me.txtFare
Me.txtUsedPortions1 = Me.txtUsedPortions
Me.txtRefundableAmount1 = Me.txtRefundableAmount
Me.txtTax1 = Me.txtTax
Me.txtASF1 = Me.txtASF
Me.txtTerminalFee1 = Me.txtTerminalFee
Me.txtInsurance1 = Me.txtInsurance
Me.txtEvat1 = Me.txtEvat
Me.txtRefundFee1 = Me.txtRefundFee
Me.txtNoShowSurcharge1 = Me.txtNoShowSurcharge
                
              Call ReCalc
        End If
       .Close
      Set RsFindTicket = Nothing
End With
End Sub

Function FindAirlineName(param) As String
Dim tmp As ADODB.Recordset
Set tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineID]=" & CDbl(param)
With tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirlineName = .Fields(1).Value
      Else
        FindAirlineName = "none"
    End If
    .Close
End With
Set tmp = Nothing
End Function


Sub ReCalc()
Dim Result As Double
Dim Credit As Double
Dim Debit(1 To 2) As Double

Result = CDbl(Me.txtFare) - CDbl(Me.txtUsedPortions)
Me.txtRefundableAmount = Format(Result, "###,##0.00")

Result = CDbl(Me.txtRefundableAmount)
Credit = CDbl(Me.txtTax) + CDbl(Me.txtASF) + CDbl(Me.txtTerminalFee) + CDbl(Me.txtInsurance) + CDbl(Me.txtEvat)
Debit(1) = CDbl(Me.txtServiceFee) + CDbl(txtRefundFee) + CDbl(Me.txtNoShowSurcharge) + CDbl(txtVoidFee)

txtRefundAmountPax = Format((Result + Credit) - Debit(1), "###,##0.00")
Result = CDbl(txtRefundAmountPax)

Debit(2) = CDbl(txtRecallComm)

Me.txtNetRefundableAmount = Format(CDbl(Result) - CDbl(Debit(2)), "###,##0.00")


Me.txtEvatAddComm = (CDbl(Me.txtCommEvatPercent) / 100) * CDbl(Me.txtRecallComm1)
Me.txtRefundableAmount1 = CDbl(Me.txtNetRefundableAmount) + CDbl(Me.txtEvatAddComm)

End Sub


Private Sub cmdCancel_Click()
Set Me.DataGrid1.DataSource = Nothing
    Me.DataGrid1.Refresh
    If Me.Check1.Value = 1 Then
           Call CancelPostRoute(Me.Check1.Caption)
    End If
    If Me.Check2.Value = 1 Then
           Call CancelPostRoute(Me.Check2.Caption)
    End If
    If Me.Check3.Value = 1 Then
           Call CancelPostRoute(Me.Check3.Caption)
    End If
    If Me.Check4.Value = 1 Then
           Call CancelPostRoute(Me.Check4.Caption)
    End If
    If Me.Check5.Value = 1 Then
           Call CancelPostRoute(Me.Check5.Caption)
    End If
        Me.Check1.Value = 0
        Me.Check2.Value = 0
        Me.Check3.Value = 0
        Me.Check4.Value = 0
        Me.Check5.Value = 0
   MsgBox "Transaction Cancelled", vbInformation
   'Me.cmdPost.Enabled = False
   Unload Me
End Sub

Private Sub cmdPost_Click()
Dim ask As Integer
ask = MsgBox("Are you sure you want to complete this refund?", vbInformation + vbYesNo)
If ask = vbYes Then
Call UpdateStatement(StatementDetail, True)
    
    If Me.Check1.Value = 1 Then
           Call PostRoute(Me.Check1.Caption, True)
    End If
    If Me.Check2.Value = 1 Then
           Call PostRoute(Me.Check2.Caption, True)
    End If
    If Me.Check3.Value = 1 Then
           Call PostRoute(Me.Check3.Caption, True)
    End If
    If Me.Check4.Value = 1 Then
           Call PostRoute(Me.Check4.Caption, True)
    End If
    If Me.Check5.Value = 1 Then
           Call PostRoute(Me.Check5.Caption, True)
    End If
    
'016805517
'SD105-102500304-1
'//===========================================================================================
'//Update SA if ticket was not paid and refunded... subtract remaining balance with amount refunded
'//
'//===========================================================================================
          
Call SA_Update(SA_Return(Me.txtTicketNo), Me.txtTicketNo, CDbl(txtNetRefundableAmount), CDbl(txtNetFare), Format(Now, "mm/dd/yyyy"), CLng(Me.Tag))
'//===========================================================================================
    
    
MsgBox "Record Save...", vbInformation
Me.Check1.Value = 0
Me.Check2.Value = 0
Me.Check3.Value = 0
Me.Check4.Value = 0
Me.Check5.Value = 0

Else
MsgBox "Cancelled...", vbInformation
End If
End Sub

Sub PostRoute(ByVal param As Long, ByVal Criteria As Boolean)
Dim Rs As New ADODB.Recordset
SQL = "UPDATE tbl_StatementTickets SET [Posted] =" & Criteria & " WHERE [StatementTicektsID]=" & CLng(param)
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans
End Sub


Sub CancelPostRoute(ByVal param As Long)
Dim Rs As New ADODB.Recordset
SQL = "UPDATE tbl_StatementTickets SET [Posted] =FALSE,[Refund]=FALSE WHERE [StatementTicektsID]=" & CLng(param)
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans
End Sub

Private Sub cmdSet_Click()
frmUserVerifyRefund.Show 1
End Sub

Private Sub Combo1_Click()
Dim OldBookMark
Dim Criteria

   OldBookMark = Me.DataGrid1.Bookmark
   Criteria = Me.DataGrid1.Columns(0).Text
    If Me.Combo1 = "Yes" Then
       If (Me.DataGrid1.Columns(10).Text) = "No" Then
            Call UpdateTicket(Criteria, True)
            Call FindTicketDetails(Criteria, "Add")
            Me.cmdPost.Enabled = True
            
          Select Case Me.DataGrid1.Row
            Case 0
                Me.Check1.Value = 1
                Me.Check1.Caption = Me.DataGrid1.Columns(0).Text
            Case 1
                Me.Check2.Value = 1
                Me.Check2.Caption = Me.DataGrid1.Columns(0).Text
            Case 2
                Me.Check3.Value = 1
                Me.Check3.Caption = Me.DataGrid1.Columns(0).Text
            Case 3
                Me.Check4.Value = 1
                Me.Check4.Caption = Me.DataGrid1.Columns(0).Text
            Case 4
                Me.Check5.Value = 1
                Me.Check5.Caption = Me.DataGrid1.Columns(0).Text
          End Select
            
       End If
    Else
        If (Me.DataGrid1.Columns(11).Text) = "Yes" Then
            MsgBox "Error this route was already refunded! You need admin rights to over-ride", vbCritical
        Else
            Call UpdateTicket(Criteria, False)
            Call FindTicketDetails(Criteria, "Subtract")
          Select Case Me.DataGrid1.Row
            Case 0
                Me.Check1.Value = 0
            Case 1
                Me.Check2.Value = 0
            Case 2
                Me.Check3.Value = 0
            Case 3
                Me.Check4.Value = 0
            Case 4
                Me.Check5.Value = 0
          End Select

        End If
    End If
    
    RsRoutes.Requery
    RsRoutes.Bookmark = OldBookMark
    Me.Combo1.Visible = False
End Sub




































Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
MoveCombo
End Sub

Private Sub Form_Load()
Me.txtRefundDate = Format(Now, "mm/dd/yyyy")
Me.DataGrid1.RowHeight = Me.Combo1.Height
Combo1.Move -10000
End Sub


Private Sub cmdFind_Click()
myPublic_AirShipID = Empty
myPublic_AccountID = Empty
Me.Option1.Value = False
LoadValues Me.txtTicketNo
Call ReCalc
End Sub


Private Sub cmdExit_Click()
If (Me.Check1.Value = 1) Or (Me.Check2.Value = 1) Or (Me.Check3.Value = 1) Or (Me.Check4.Value = 1) Or (Me.Check5.Value = 1) Then
    MsgBox "Record was changed click Post to continue!", vbInformation
    Me.cmdPost.SetFocus
Else
    Unload Me
End If
End Sub

Private Sub cmdFind_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub Text1_Change()

End Sub

Private Sub txtASF_GotFocus()
With Me.txtASF
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With

End Sub


Private Sub txtCommEvatPercent_Change()
ReCalc
End Sub

Private Sub txtInsurance_GotFocus()
With Me.txtInsurance
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With

End Sub

Private Sub txtNoShowSurcharge_Change()
On Error GoTo ErrExit
ReCalc
Exit Sub
ErrExit:
txtNoShowSurcharge = "0.00"
Call kulotHL(txtNoShowSurcharge)
End Sub

Private Sub txtNoShowSurcharge_GotFocus()
With Me.txtNoShowSurcharge
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With

End Sub

Private Sub txtNoShowSurcharge_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtRecallComm_Change()
ReCalc
End Sub

Private Sub txtRecallComm_GotFocus()
With Me.txtRecallComm
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With
End Sub

Private Sub txtRecallComm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cmdPost.SetFocus
End If
End Sub

Private Sub txtRefundFee_Change()
On Error GoTo ErrExit
ReCalc
Exit Sub
ErrExit:
txtRefundFee = "0.00"
Call kulotHL(txtRefundFee)
End Sub

Private Sub txtRefundFee_LostFocus()
txtRefundFee.Enabled = False
End Sub

Private Sub txtServiceFee_Change()
On Error GoTo ErrExit
ReCalc
Exit Sub
ErrExit:
txtServiceFee = "0.00"
Call kulotHL(txtServiceFee)
End Sub




Private Sub txtServiceFee_GotFocus()
With Me.txtServiceFee
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With

End Sub

Private Sub txtServiceFee_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtTax_Change()
ReCalc
End Sub

Private Sub txtTax_GotFocus()
With Me.txtTax
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With
End Sub


Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtTerminalFee_GotFocus()
With Me.txtTerminalFee
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With

End Sub

Private Sub txtTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtUsedPortions_Change()
ReCalc
End Sub

Private Sub txtUsedPortions_GotFocus()
With Me.txtUsedPortions
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With
End Sub

Private Sub txtUsedPortions_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtVoidFee_Change()
On Error GoTo ErrExit
ReCalc
Exit Sub
ErrExit:
txtVoidFee = "0.00"
Call kulotHL(txtVoidFee)
End Sub


Private Sub MoveCombo()
   'On Error GoTo Error_Handler
    Combo1.Visible = True
    Me.DataGrid1.Refresh
        Combo1.Move 4320, _
            DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row), 1230
        'Combo1.ZOrder
        'Combo1.SetFocus
        Combo1.Text = Me.DataGrid1.Columns(10).Text
        Exit Sub
Error_Handler:
    Combo1.Move -10000
    If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub


