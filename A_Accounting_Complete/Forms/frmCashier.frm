VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmCashier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cashier"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14730
   ControlBox      =   0   'False
   Icon            =   "frmCashier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Accounting.ucGradContainer ucGradContainer1 
      Height          =   9675
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   14730
      _extentx        =   25982
      _extenty        =   17066
      headericon      =   "frmCashier.frx":6852
      captionfont     =   "frmCashier.frx":D0B6
      iconsize        =   0
      backcolor2      =   16244947
      backcolor1      =   16244947
      caption         =   "         Cashier Payments"
      borderwidth     =   2
      Begin VB.TextBox txtOR 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11640
         TabIndex        =   51
         Top             =   480
         Width           =   2865
      End
      Begin MSComctlLib.ImageList SmallImages 
         Left            =   3255
         Top             =   45
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
               Picture         =   "frmCashier.frx":D0E4
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":DDBE
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":EC10
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":FA62
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1033C
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":10C16
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":114F0
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":11EBA
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":12794
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":12AAE
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":13388
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":13C62
               Key             =   "IMG12"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1453C
               Key             =   "IMG13"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":14856
               Key             =   "IMG14"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":15130
               Key             =   "IMG15"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":15A0A
               Key             =   "IMG16"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":162E4
               Key             =   "IMG17"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":16BBE
               Key             =   "IMG18"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":17498
               Key             =   "IMG19"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":17D72
               Key             =   "IMG20"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1864C
               Key             =   "IMG21"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":18F26
               Key             =   "IMG22"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":19800
               Key             =   "IMG23"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1A0DA
               Key             =   "IMG24"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1A9B4
               Key             =   "IMG25"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1B28E
               Key             =   "IMG26"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1BB68
               Key             =   "IMG27"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1C442
               Key             =   "IMG28"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1CD1C
               Key             =   "IMG29"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1D5F6
               Key             =   "IMG30"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1DEAC
               Key             =   "IMG31"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1E786
               Key             =   "IMG32"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1EBD8
               Key             =   "IMG33"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":1F02A
               Key             =   "IMG34"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":217DC
               Key             =   "IMG35"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashier.frx":22A5E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   2130
         Left            =   6945
         TabIndex        =   38
         Top             =   885
         Width           =   7740
         Begin VB.TextBox txtTotal 
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
            Height          =   420
            Left            =   4710
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   630
            Width           =   2865
         End
         Begin VB.TextBox txtMisc 
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
            ForeColor       =   &H000000C0&
            Height          =   420
            Left            =   4710
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   210
            Width           =   2865
         End
         Begin VB.TextBox txtAmountBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   495
            Left            =   4710
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   1515
            Width           =   2880
         End
         Begin VB.TextBox txtAmountTendered 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   4710
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   1050
            Width           =   2880
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3555
            TabIndex        =   46
            Top             =   705
            Width           =   930
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MISC :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   345
            Left            =   3780
            TabIndex        =   45
            Top             =   225
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT BALANCE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   1890
            TabIndex        =   44
            Top             =   1680
            Width           =   2595
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT TENDERED :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   1740
            TabIndex        =   43
            Top             =   1155
            Width           =   2745
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Statement Details"
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
         Height          =   1635
         Left            =   3465
         TabIndex        =   28
         Top             =   5685
         Visible         =   0   'False
         Width           =   7560
         Begin VB.TextBox Text2 
            Height          =   300
            Left            =   4935
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   1155
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtAirline 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2475
            TabIndex        =   32
            Top             =   1140
            Width           =   2355
         End
         Begin VB.TextBox txtStatementNo 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   31
            Top             =   1140
            Width           =   2355
         End
         Begin VB.TextBox txtStatementDate 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   30
            Top             =   510
            Width           =   2340
         End
         Begin VB.TextBox txtStatementType 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4635
            TabIndex        =   29
            Top             =   510
            Width           =   2205
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AIRLINE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2490
            TabIndex        =   37
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STATEMENT NO. :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   36
            Top             =   915
            Width           =   1395
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STATEMENT DATE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   35
            Top             =   285
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STATEMENT TYPE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4620
            TabIndex        =   34
            Top             =   270
            Width           =   1515
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1530
         Left            =   45
         TabIndex        =   15
         Top             =   855
         Width           =   6885
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   975
            Left            =   2400
            ScaleHeight     =   915
            ScaleWidth      =   4110
            TabIndex        =   18
            Top             =   360
            Width           =   4170
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   75
               TabIndex        =   19
               Top             =   360
               Width           =   2685
            End
            Begin LVbuttons.LaVolpeButton cmdFind 
               Height          =   480
               Left            =   2835
               TabIndex        =   20
               Top             =   360
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
               MICON           =   "frmCashier.frx":25210
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Statement Number:"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   105
               Width           =   2640
            End
         End
         Begin VB.OptionButton OptSearchSBS 
            Caption         =   "Search by Statement #"
            Height          =   345
            Left            =   165
            TabIndex        =   17
            Top             =   480
            Value           =   -1  'True
            Width           =   2115
         End
         Begin VB.OptionButton OptSearchSBN 
            Caption         =   "Search by Name"
            Height          =   345
            Left            =   165
            TabIndex        =   16
            Top             =   900
            Width           =   2115
         End
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7E0D3&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1635
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   7905
         Width           =   1890
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Customer Information"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   6990
         TabIndex        =   9
         Top             =   8505
         Width           =   7680
         Begin VB.TextBox txtReceivedFrom 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   255
            TabIndex        =   11
            Top             =   465
            Width           =   2670
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2970
            TabIndex        =   10
            Top             =   465
            Width           =   4665
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RECEIVED FROM :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   270
            TabIndex        =   13
            Top             =   195
            Width           =   1635
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ADDRESS :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2970
            TabIndex        =   12
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1695
         Left            =   45
         ScaleHeight     =   1635
         ScaleWidth      =   6825
         TabIndex        =   5
         Top             =   2415
         Width           =   6885
         Begin VB.Frame Frame5 
            Caption         =   "Payment Mode"
            Height          =   1110
            Left            =   60
            TabIndex        =   53
            Top             =   465
            Width           =   6735
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   4935
               TabIndex        =   57
               Text            =   "0.00"
               Top             =   630
               Width           =   1665
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Set Dollar Rate"
               Height          =   525
               Left            =   180
               TabIndex        =   56
               Top             =   360
               Width           =   1740
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Dollar"
               Height          =   435
               Left            =   4455
               TabIndex        =   55
               Top             =   135
               Width           =   2130
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Peso"
               Height          =   390
               Left            =   2295
               TabIndex        =   54
               Top             =   135
               Value           =   -1  'True
               Width           =   1830
            End
            Begin VB.Label Label13 
               Caption         =   "Current Peso Dollar Exchange Rate :"
               Height          =   225
               Left            =   2265
               TabIndex        =   58
               Top             =   615
               Width           =   2640
            End
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Domestic Payment"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   300
            TabIndex        =   8
            Top             =   60
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "International Payment"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2370
            TabIndex        =   7
            Top             =   60
            Width           =   1920
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Documentation Payment"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4530
            TabIndex        =   6
            Top             =   60
            Width           =   2115
         End
      End
      Begin VB.TextBox txtTotCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7E0D3&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5940
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   8010
         Width           =   1605
      End
      Begin VB.TextBox txtTotCard 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7E0D3&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7590
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   8010
         Width           =   1605
      End
      Begin VB.TextBox txtTotCheck 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7E0D3&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9225
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   8010
         Width           =   1605
      End
      Begin VB.TextBox txtTotOthers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7E0D3&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10875
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   8010
         Width           =   1590
      End
      Begin LVbuttons.LaVolpeButton cmdNew 
         Height          =   720
         Left            =   120
         TabIndex        =   22
         Top             =   8655
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1270
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
         MICON           =   "frmCashier.frx":2522C
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
         Height          =   720
         Left            =   1860
         TabIndex        =   23
         Top             =   8655
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1270
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
         MICON           =   "frmCashier.frx":25248
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
         Height          =   720
         Left            =   3525
         TabIndex        =   24
         Top             =   8655
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   1270
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
         MICON           =   "frmCashier.frx":25264
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
         Height          =   720
         Left            =   5250
         TabIndex        =   25
         Top             =   8655
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1270
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
         MICON           =   "frmCashier.frx":25280
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   3570
         Left            =   45
         TabIndex        =   26
         Top             =   4155
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   6297
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   28
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Statement #"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Airline"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Airline ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cash"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Check"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Credit Card"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Other"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Check Number"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Check Bank"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Check Branch"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Card Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Card Number"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Card Holder"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Date"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Check Date"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Post Dated"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "No of Days"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "WhichAcc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "BankAcc1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "BankAcc2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "BankAcc3"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "BankAcc4"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Card Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "CardExpireMon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "CardExpireYear"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "CardSecureCode"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFFICIAL RECEIPT # :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9645
         TabIndex        =   52
         Top             =   480
         Width           =   1770
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total Others"
         Height          =   180
         Left            =   10905
         TabIndex        =   50
         Top             =   7800
         Width           =   1605
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total Card"
         Height          =   270
         Left            =   9255
         TabIndex        =   49
         Top             =   7800
         Width           =   1650
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total Check"
         Height          =   270
         Left            =   7620
         TabIndex        =   48
         Top             =   7800
         Width           =   1770
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total Cash"
         Height          =   225
         Left            =   6000
         TabIndex        =   47
         Top             =   7785
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-TOTAL :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   135
         TabIndex        =   27
         Top             =   7890
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Selection As String
Dim RsGrid As ADODB.Recordset
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim i As Integer
Dim mylist As ListItem

'Call ClearMe
Set Rs = New ADODB.Recordset

If Me.Option1 Then
    SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & Me.Text1 & "'"
End If

If Me.Option2 Then
    SQL = "SELECT * FROM tbl_Statement_INTL WHERE [SAno]='" & Me.Text1 & "'"
End If

If Me.Option3 Then
    SQL = "SELECT * FROM tbl_PPT WHERE [StatementNo]='" & Me.Text1 & "'"
End If

With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            If Me.Option1 Then
               If Rs.Fields("Paid").Value Then
                    MsgBox "This statement already Paid!", vbInformation
                    Call kulotHL(Me.Text1)
                    Exit Sub
               Else
                          
                Me.txtStatementDate = Rs.Fields("Date").Value
                txtStatementType.Text = Me.Option1.Caption
                Me.Text2.Text = Rs.Fields("Airline").Value
                Me.txtAirline = FindAirlineName(Me.Text2)
                
        
                If Me.ListView2.ListItems.Count > 0 Then
                '//----------------------------------------------------------------------------
                '// Loop used to check for duplicate SA
                '//----------------------------------------------------------------------------
                
                        For i = 1 To Me.ListView2.ListItems.Count
                                If Me.ListView2.ListItems(i).Text = Rs.Fields("sNumber").Value Then
                                    Exit Sub
                                End If
                        Next i
                 '//----------------------------------------------------------------------------
                Set mylist = ListView2.ListItems.Add(, , Rs.Fields("sNumber").Value)
                            mylist.SubItems(1) = Format(Rs.Fields("Balance").Value, "###,##0.00")
                            mylist.SubItems(2) = FindAirlineName(Rs.Fields("Airline").Value)
                            mylist.SubItems(3) = Rs.Fields("Airline").Value
                            mylist.SubItems(4) = "0.00"
                            mylist.SubItems(5) = "0.00"
                            mylist.SubItems(6) = "0.00"
                            mylist.SubItems(7) = "0.00"
                Else
                Set mylist = ListView2.ListItems.Add(, , Rs.Fields("sNumber").Value)
                            mylist.SubItems(1) = Format(Rs.Fields("Balance").Value, "###,##0.00")
                            mylist.SubItems(2) = FindAirlineName(Rs.Fields("Airline").Value)
                            mylist.SubItems(3) = Rs.Fields("Airline").Value
                            mylist.SubItems(4) = "0.00"
                            mylist.SubItems(5) = "0.00"
                            mylist.SubItems(6) = "0.00"
                            mylist.SubItems(7) = "0.00"
                End If
                            txtSubtotal.Text = Format(SumListView(), "###,##0.00")
                End If
                
            End If

            '===========================================================
            If Me.Option2 Then
               If Rs.Fields("Paid").Value Then
                    MsgBox "This statement already Paid!", vbInformation
                    Call kulotHL(Me.Text1)
                    Exit Sub
               Else
                    Me.txtStatementDate = Rs.Fields("SA Date").Value
                    txtStatementType.Text = Me.Option2.Caption
                    txtSubtotal.Text = Format(Rs.Fields("Balance").Value, "###,##0.00")
                    'Me.Text2.Text = Rs.Fields("AirlineID").Value
                    'Me.txtairline = FindAirlineName(Me.Text2)
                    
                If Me.ListView2.ListItems.Count > 0 Then
                '//----------------------------------------------------------------------------
                '// Loop used to check for duplicate SA
                '//----------------------------------------------------------------------------
                
                        For i = 1 To Me.ListView2.ListItems.Count
                                If Me.ListView2.ListItems(i).Text = Rs.Fields("SAno").Value Then
                                    Exit Sub
                                End If
                        Next i
                 '//----------------------------------------------------------------------------
                Set mylist = ListView2.ListItems.Add(, , Rs.Fields("SAno").Value)
                            mylist.SubItems(1) = Format(Rs.Fields("Balance").Value, "###,##0.00")
                            'mylist.SubItems(2) = FindAirlineName(Rs.Fields("Airline").Value)
                            'mylist.SubItems(3) = Rs.Fields("Airline").Value
                            mylist.SubItems(4) = "0.00"
                            mylist.SubItems(5) = "0.00"
                            mylist.SubItems(6) = "0.00"
                            mylist.SubItems(7) = "0.00"
                Else
                Set mylist = ListView2.ListItems.Add(, , Rs.Fields("SAno").Value)
                            mylist.SubItems(1) = Format(Rs.Fields("Balance").Value, "###,##0.00")
                            'mylist.SubItems(2) = FindAirlineName(Rs.Fields("Airline").Value)
                            'mylist.SubItems(3) = Rs.Fields("Airline").Value
                            mylist.SubItems(4) = "0.00"
                            mylist.SubItems(5) = "0.00"
                            mylist.SubItems(6) = "0.00"
                            mylist.SubItems(7) = "0.00"
                End If
                            txtSubtotal.Text = Format(SumListView(), "###,##0.00")
                End If
                
            End If
            
            '===========================================================
            If Me.Option3 Then
             If Rs.Fields("Paid").Value Then
                    MsgBox "This statement already Paid!", vbInformation
                    Call kulotHL(Me.Text1)
                        Exit Sub
                Else
                    Me.txtStatementDate = Rs.Fields("Date").Value
                    txtStatementType.Text = Me.Option3.Caption
                    txtSubtotal.Text = Format(Rs.Fields("Balance").Value, "###,##0.00")
                    Me.Text2.Text = FindAirline
                    Me.txtAirline = "NONE"
             End If
            End If
            '===========================================================
            Call ReCompute
                txtStatementNo.Text = Me.Text1
            Call kulotHL(Me.Text1)
        Else
                MsgBox "No match!!!", vbCritical
'            Call ClearMe
            Call kulotHL(Me.Text1)
        End If
                .Close
                Set Rs = Nothing
End With
End Sub

Function SumListView() As Double
Dim Y As Integer
Dim Tmp As Double

If Me.ListView2.ListItems.Count > 0 Then
Me.Caption = Me.ListView2.ListItems.Count
        For Y = 1 To Me.ListView2.ListItems.Count
            If Me.ListView2.ListItems(Y).SubItems(1) <> "" Then
                Tmp = Tmp + CDbl(Me.ListView2.ListItems(Y).SubItems(1))
            End If
        Next Y
Else
        Tmp = 0
End If

SumListView = Tmp

End Function


Private Sub cmdNew_Click()
txtOR = AutoIncrement(Get_OR_LastNum)
Call ClearMe
Me.ListView2.ListItems.Clear
Me.Text1 = ""
Me.txtTotal = "0.00"
Me.txtSubtotal = "0.00"
Me.txtAmountTendered = "0.00"
Me.Text1.SetFocus
End Sub
Sub ClearMe()
Me.Frame1.Enabled = True


'Me.Text1 = Empty
Me.txtStatementType = Empty
txtStatementNo = Empty
'txtSubtotal = "0.00"
txtMisc = "0.00"
txtTotal = "0.00"
txtAmountTendered = "0.00"
txtAmountBalance = "0.00"
txtChange = "0.00"
txtCash = "0.00"
Me.txtSubtotal = "0.00"

Me.Text2 = Empty
Me.txtAirline = Empty


End Sub

Function FindAirlineName(Param) As String
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineID]=" & CDbl(Param)
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirlineName = .Fields(1).Value
      Else
        FindAirlineName = "none"
    End If
    .Close
End With
Set Tmp = Nothing
End Function


Function FindAirline() As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='NONE'"
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


Private Sub cmdOverRide_Click()
'Set Me.DataGrid1.DataSource = Nothing
End Sub

Sub AddtoCurrentAccNo(Param, amt)
Dim RsAccNo As New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & Param & "'"
With RsAccNo
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                .Fields("Current Balance").Value = .Fields("Current Balance").Value + CDbl(amt)
                .Update
                .MoveNext
            Loop
        End If
        .Close
End With

End Sub




Function Is_Paid(Opt As OptionButton, nStr As String) As Boolean
If Opt.Value = True Then
  If CheckIfPaid(nStr) Then
        Is_Paid = True
        MsgBox "This statement is already fully paid!", vbInformation
      Else
        Is_Paid = False
  End If
End If
End Function

Function Is_BelowZero(Opt As OptionButton, nText As TextBox) As Boolean
If Opt.Value = True Then
  If CDbl(nText) <= 0 Then
        Is_BelowZero = True
        MsgBox "Amount should not be less than or equal to zero", vbCritical
        Call kulotHL(nText)
      Else
        Is_BelowZero = False
  End If
End If

End Function



Sub UpdatePassbook(ByVal strSa As String, ByVal nAmt As Double, ByVal CheckNo As String, ByVal CheckDate As String, Optional Desc As String, Optional ByVal AccNo As String, Optional ORno As String, Optional strAir As String, Optional nCash, Optional nCard, Optional nCheck, Optional nOthers, Optional nCardName, Optional nCardNumber, Optional nCardHolder, Optional nBank1, Optional nBank2, Optional nBank3, Optional nBank4, Optional nPDC As Boolean, Optional c1, Optional c2, Optional c3, Optional c4)
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

'MsgBox "Balance :" & TempBal & " Amount :" & nAmt

SQL = "SELECT * FROM tbl_BankPassbook"
With RsPassbk
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("Deposit Date").Value = Format(Now, "mm/dd/yyyy")
        .Fields("SA no").Value = strSa
        .Fields("Check No").Value = CheckNo
        .Fields("Check Date").Value = CheckDate
        .Fields("Voucher No").Value = "n/a"
        .Fields("Description").Value = Desc
        .Fields("Credit").Value = nAmt
        .Fields("Debit").Value = 0
        .Fields("Account Number").Value = AccNo
        .Fields("Balance").Value = TempBal + nAmt
        .Fields("Cash Amount").Value = CDbl(nCash)
        .Fields("Card Amount").Value = CDbl(nCard)
        .Fields("Check Amount").Value = CDbl(nCheck)
        .Fields("Others Amount").Value = CDbl(nOthers)
        .Fields("ORno").Value = ORno
       If Not Me.Option2 Then
        .Fields("Airline").Value = strAir
       End If
        .Fields("Card Name").Value = nCardName
        .Fields("Card Number").Value = nCardNumber
        .Fields("Card Holder").Value = nCardHolder
        .Fields("Bank1").Value = nBank1
        .Fields("Bank2").Value = nBank2
        .Fields("Bank3").Value = nBank3
        .Fields("Bank4").Value = nBank4
        .Fields("Post Dated").Value = nPDC
        
        
        .Fields("CardAdr").Value = c1
        .Fields("CardExpireMon").Value = c2
        .Fields("CardExpireYear").Value = c3
        .Fields("CardSecureCode").Value = c4
        
        .Update
End With

SQL = "UPDATE tbl_AccountsSetting SET [Current Balance] = " & _
              CDbl(TempBal + nAmt) & " WHERE [Account Number]= '" & UCase(AccNo) & "'"
              cn.BeginTrans
                    cn.Execute SQL
              cn.CommitTrans
Exit Sub
FailSafe_Error:
cn.RollbackTrans
End Sub

Sub SaveToBank(Param, SettingsID)
Dim RsAcc As New ADODB.Recordset
SQL = "SELECT * FROM tbl_BankAccounts"
With RsAcc
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        '.AddNew
        '.Fields(1).Value = CDbl(SettingsID)
        '.Fields(2).Value = CDbl(Me.txtAmountTendered)
        '.Fields(3).Value = Format(Now, "mm/dd/yyyy")
        '.Fields(4).Value = Param
        '.Fields(5).Value = Me.txtStatementDate
        '.Fields(6).Value = IIf(Me.OptCard, Me.OptCard.Caption, IIf(Me.optCash, Me.optCash.Caption, Me.OptCheck.Caption))
        '.Fields(7).Value = CDbl(Me.Text2)
        '.Fields(8).Value = Me.txtAirline
        '.Update
        .Close
End With
Exit Sub
End Sub


Function ReturnAccNo(Param) As String
Dim RsAccNo As New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [AirlineId]=" & Param
With RsAccNo
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            ReturnAccNo = .Fields(3).Value
        End If
        .Close
End With

End Function


Function ReturnSettingsID(Param) As Long
Dim RsAccNo As New ADODB.Recordset
SQL = "SELECT * FROM tbl_AccountsSetting WHERE [AirlineId]=" & Param
With RsAccNo
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            ReturnSettingsID = .Fields(0).Value
        End If
        .Close
End With

End Function

Private Sub cmdPost_Click()
'On Error GoTo FailSafe_Error

Dim RsMyOR                  As ADODB.Recordset
Dim RsPayment               As ADODB.Recordset
Dim RsPaymentDetails        As ADODB.Recordset
Dim RsStatement             As ADODB.Recordset
Dim RsMyTmp                 As New ADODB.Recordset

Dim SQL                     As String
Dim ask                     As Integer
Dim myTempORid              As Long
Dim StatementCtr            As Integer
Dim myRpt                   As New RptOR
Dim i, j                    As Integer


    Set RsPayment = New ADODB.Recordset
    Set RsPaymentDetails = New ADODB.Recordset
    Set RsMyOR = New ADODB.Recordset
    
    StatementCtr = 0
    
    If Is_Paid(Me.Option1, "Domestic") Then: Exit Sub
    If Is_Paid(Me.Option2, "International") Then: Exit Sub
    If Is_Paid(Me.Option3, "Documents") Then: Exit Sub

    
    If CheckNull(Me.txtReceivedFrom) Then
            MsgBox "Please supply received from!", vbInformation
            Call kulotHL(Me.txtReceivedFrom)
            Exit Sub
    End If
    If CheckNull(Me.txtAddress) Then
            MsgBox "Please supply address!", vbInformation
            Call kulotHL(Me.txtAddress)
            Exit Sub
    End If
    
    ask = MsgBox("Are the entry correct?", vbInformation + vbYesNo)
    If ask = vbYes Then
        Call ReCompute
        cn.BeginTrans


SQL = "SELECT * FROM tbl_OR"
With RsMyOR
            '.CursorLocation = adUseClient
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If Not FindDupOR(Me.txtOR) Then
                .AddNew
                RsMyOR("ORno").Value = Me.txtOR
                RsMyOR("Amount").Value = CDbl(Me.txtAmountTendered)
                RsMyOR("Date").Value = Format(Now, "mm/dd/yyyy")
                RsMyOR("Void").Value = False
                .Update
                myTempORid = RsMyOR("OrID").Value               '<=====
            End If
End With

SQL = "SELECT * FROM tbl_CashierPayments"
With RsPayment
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            
    For i = 1 To Me.ListView2.ListItems.Count
                .AddNew
                .Fields("OrID").Value = myTempORid              '<=====
                .Fields("Branch Number").Value = WhichBranch.Fields(2).Value
                .Fields("Statement Type").Value = IIf(Me.Option1.Value, Me.Option1.Caption, IIf(Me.Option2, Me.Option2.Caption, Me.Option3.Caption))
                .Fields("Statement Number").Value = Me.ListView2.ListItems(i).Text
                .Fields("Amount Due").Value = Me.ListView2.ListItems(i).SubItems(1)
                .Fields("Balance").Value = Me.ListView2.ListItems(i).SubItems(18)
                .Fields("Date").Value = txtStatementDate
                .Fields("Void").Value = False
                .Update
                Call MarkStatement(Me.ListView2.ListItems(i).Text, Me.ListView2.ListItems(i).SubItems(18), i)
    Next i
    
End With


'//--------------------------------------------------------------------
'//UpdatePassbook(strSa,nAmt,CheckNo,CheckDate,Desc,AccNo,ORno,strAir)
'//--------------------------------------------------------------------
Dim myTotalCredit   As Double
Dim myCash          As Double
Dim myCard          As Double
Dim myCheck         As Double
Dim myOthers        As Double
Dim myCardHolder
Dim myCardName
Dim myCardNumber

Dim myCheckNum          As String
Dim myCheckBank         As String
Dim myCheckBranch       As String
Dim myTmpAccno          As String
Dim myLilyHoAccno       As String
Dim myPAL_Accno         As String
Dim myPALHSBC_Accno     As String
Dim myTemp_Accno        As String
Dim myBank1           As String
Dim myBank2           As String
Dim myBank3           As String
Dim myBank4           As String
Dim mynPDC            As Boolean

Dim myC1, myC2, myC3, myC4

            For i = 1 To Me.ListView2.ListItems.Count
           
            myCash = CDbl(Me.ListView2.ListItems(i).SubItems(4))
            myCheck = CDbl(Me.ListView2.ListItems(i).SubItems(5))
            myCard = CDbl(Me.ListView2.ListItems(i).SubItems(6))
            myOthers = CDbl(Me.ListView2.ListItems(i).SubItems(7))
             
            myCardName = Me.ListView2.ListItems(i).SubItems(11)
            myCardNumber = Me.ListView2.ListItems(i).SubItems(12)
            myCardHolder = Me.ListView2.ListItems(i).SubItems(13)
            
            mynPDC = IIf(Me.ListView2.ListItems(i).SubItems(16) = "Yes", True, False)
            
            
            
'//========================================================================================
'//         getGlobal LilyHo Accno
'//========================================================================================
            
            myLilyHoAccno = myGlobal_LilyHo_AccNo
            
            myBankAcc1 = Me.ListView2.ListItems(i).SubItems(20)
            myBankAcc2 = Me.ListView2.ListItems(i).SubItems(21)
            myBankAcc3 = Me.ListView2.ListItems(i).SubItems(22)
            myBankAcc4 = Me.ListView2.ListItems(i).SubItems(23)
            
            
            myC1 = Me.ListView2.ListItems(i).SubItems(24)
            myC2 = Me.ListView2.ListItems(i).SubItems(25)
            myC3 = Me.ListView2.ListItems(i).SubItems(26)
            myC4 = Me.ListView2.ListItems(i).SubItems(27)
                        
            myTotalCredit = myCash + myCard + myCheck + myOthers
            
            If myCash > 0 Then
                    Call UpdatePassbook(Me.ListView2.ListItems(i).Text, _
                                        myCash, _
                                        "", _
                                        "", _
                                        "Cash", _
                                        myLilyHoAccno, _
                                        Me.txtOR, _
                                        Me.ListView2.ListItems(i).SubItems(3), _
                                        myCash, _
                                        0, _
                                        0, _
                                        0, _
                                        "", _
                                        "", _
                                        "", _
                                        myBankAcc1, _
                                        myBankAcc2, _
                                        myBankAcc3, _
                                        myBankAcc4, mynPDC, myC1, myC2, myC3, myC4)

            End If
                        
            
            If myCard > 0 Then
                    
                myTmpAccno = Me.ListView2.ListItems(i).SubItems(19)
            
                    Call UpdatePassbook(Me.ListView2.ListItems(i).Text, _
                                        myCard, _
                                        "", _
                                        "", _
                                        "Card", _
                                        myTmpAccno, _
                                        Me.txtOR, _
                                        Me.ListView2.ListItems(i).SubItems(3), _
                                        0, _
                                        myCard, _
                                        0, _
                                        0, _
                                        myCardName, _
                                        myCardNumber, _
                                        myCardHolder, _
                                        myBankAcc1, _
                                        myBankAcc2, _
                                        myBankAcc3, _
                                        myBankAcc4, mynPDC, myC1, myC2, myC3, myC4)
            End If
            
            If myCheck > 0 Then
                    Call UpdatePassbook(Me.ListView2.ListItems(i).Text, _
                                        myCheck, _
                                        Me.ListView2.ListItems(i).SubItems(8), _
                                        Me.ListView2.ListItems(i).SubItems(15), _
                                        "Check", _
                                        myLilyHoAccno, _
                                        Me.txtOR, _
                                        Me.ListView2.ListItems(i).SubItems(3), _
                                        0, _
                                        0, _
                                        myCheck, _
                                        0, _
                                        "", _
                                        "", _
                                        "", _
                                        myBankAcc1, _
                                        myBankAcc2, _
                                        myBankAcc3, _
                                        myBankAcc4, mynPDC, myC1, myC2, myC3, myC4)
            End If
            
            If myOthers > 0 Then
                    Call UpdatePassbook(Me.ListView2.ListItems(i).Text, _
                                        myOthers, _
                                        "", _
                                        "", _
                                        "Others", _
                                        myLilyHoAccno, _
                                        Me.txtOR, _
                                        Me.ListView2.ListItems(i).SubItems(3), _
                                        0, _
                                        0, _
                                        0, _
                                        myOthers, _
                                        "", _
                                        "", _
                                        "", _
                                        myBankAcc1, _
                                        myBankAcc2, _
                                        myBankAcc3, _
                                        myBankAcc4, mynPDC, myC1, myC2, myC3, myC4)
            End If
            
                    
            Next i
    
            cn.CommitTrans
            MsgBox "Record Save...", vbInformation
          '  With myRpt
          '            .DataControl1.Connection = cn
          '            .DataControl1.Source = "SELECT * FROM qryOR WHERE [ORno]='" & Me.txtOR & "'"
          '            .Show 1
          '  End With
End If

Exit Sub

FailSafe_Error:
cn.RollbackTrans
Select Case Err.Number
Case -2147467259
MsgBox "BANK ACCOUNT NOT SET!!!! " & Chr(13) & _
        "Please set first the bank account of this Airline/Shipping Line", vbCritical
Case Else
MsgBox "There was an Error in saving the record"
End Select

End Sub

Function CheckOR(ByVal strSa As String, ByVal strAirline) As Boolean
Dim rsOR As New ADODB.Recordset
SQL = "SELECT * FROM tbl_BankPassbook WHERE [ORno]='" & strSa & "' AND [Airline]='" & strAirline & "'"
With rsOR
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
       If .RecordCount > 0 Then
            CheckOR = True
        Else
            CheckOR = False
       End If
     .Close
Set rsOR = Nothing
End With
End Function


Sub MarkStatement(ByVal strStatement As String, ByVal cBal As Double, ByVal pIndex As Integer)
Dim RsStatement As New ADODB.Recordset
Dim SQL As String
Dim i As Integer

' For i = 1 To Me.ListView2.ListItems.Count
 
' Next i

'For Domestic
'=====================
If Me.Option1 Then
SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & strStatement & "'"
Set RsStatement = New ADODB.Recordset
With RsStatement
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               .Fields("Down").Value = True
               
               If CDbl(Me.ListView2.ListItems(pIndex).SubItems(18)) = 0 Then
                .Fields("Paid").Value = True
                .Fields("Down").Value = False
               Else
                .Fields("Paid").Value = False
               End If
                
                .Fields("Credit Card Activated").Value = True

                .Fields("Balance").Value = cBal
                .Update
            End If


End With
End If

'for international
'====================
If Me.Option2 Then
SQL = "SELECT * FROM tbl_Statement_INTL WHERE [SANo]='" & strStatement & "'"
Set RsStatement = New ADODB.Recordset
With RsStatement
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .Fields("Down").Value = True
               If CDbl(Me.txtAmountBalance) = 0 Then
                .Fields("Paid").Value = True
                .Fields("Down").Value = False
               Else
                .Fields("Paid").Value = False
               End If
               

                .Fields("Credit Card Activated").Value = True

                .Fields("Balance").Value = CDbl(Me.txtAmountBalance)
                .Update
            End If


End With
End If

'for passpoting
'========================
If Me.Option3 Then
SQL = "SELECT * FROM tbl_PPT WHERE [StatementNo]='" & strStatement & "'"
Set RsStatement = New ADODB.Recordset
With RsStatement
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
            .Fields("Down").Value = True
               If CDbl(Me.txtAmountBalance) = 0 Then
                .Fields("Paid").Value = True
                .Fields("Down").Value = False
               Else
                .Fields("Paid").Value = False
               End If
               
               
                .Fields("Credit Card Activated").Value = True
               
                .Fields("Balance").Value = CDbl(Me.txtAmountBalance)
                .Update
            End If

End With
End If

End Sub

Function FindDupOR(ByVal kjOR As String) As Boolean
Dim rsOR As ADODB.Recordset
SQL = "SELECT * FROM tbl_OR WHERE [ORno]='" & kjOR & "'"
Set rsOR = New ADODB.Recordset
With rsOR
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FindDupOR = True
        Else
            FindDupOR = False
        End If
        .Close
     Set rsOR = Nothing
End With
End Function





Private Sub Form_Load()
myGlobal_LilyHo_AccNo = "9999999999999"
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Item.Checked = True Then
    frmCashierPaymentOpt.Tag = Item.Index
    frmCashierPaymentOpt.Show 1
Else
   ' Item.SubItems(4) = "0.00"
   ' Item.SubItems(5) = "0.00"
   ' Item.SubItems(6) = "0.00"
   ' Item.SubItems(7) = "0.00"
    Call ReCompute
End If
End Sub



Private Sub Option1_Click()
Call kulotHL(Me.Text1)
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 Call kulotHL(Me.Text1)
End If
End Sub

Private Sub Option2_Click()
Call kulotHL(Me.Text1)
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call kulotHL(Me.Text1)
End If

End Sub

Private Sub Option3_Click()
Call kulotHL(Me.Text1)
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call kulotHL(Me.Text1)
End If
End Sub

Private Sub OptSearchSBN_Click()
Me.Text1 = ""
frmSearchName.Show 1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.Text1 = UCase(Me.Text1)
    Call cmdFind_Click
End If
End Sub


Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Function TotalPayments() As Double
Dim i As Integer
Dim ctr As Integer
Static TmpAmount As Double

TmpAmount = 0
If Me.ListView2.ListItems.Count > 0 Then
    ctr = Me.ListView2.ListItems.Count
    For i = 1 To ctr Step 1
        TmpAmount = TmpAmount + _
            CDbl(Me.ListView2.ListItems.Item(i).SubItems(4)) + _
            CDbl(Me.ListView2.ListItems.Item(i).SubItems(5)) + _
            CDbl(Me.ListView2.ListItems.Item(i).SubItems(6)) + _
            CDbl(Me.ListView2.ListItems.Item(i).SubItems(7))
    Next i
    TotalPayments = TmpAmount
Else
    TotalPayments = 0
End If
End Function


Function SumMyList(ByVal myIndex As Long) As Double
Dim Y As Integer
Dim Tmp As Double

If Me.ListView2.ListItems.Count > 0 Then
        For Y = 1 To Me.ListView2.ListItems.Count
            If Me.ListView2.ListItems(Y).SubItems(myIndex) <> "" Then
                Tmp = Tmp + CDbl(Me.ListView2.ListItems(Y).SubItems(myIndex))
            End If
        Next Y
Else
        Tmp = 0
End If

SumMyList = Tmp

End Function


Sub ReCompute()
Dim cTot    As Double
Dim cBal    As Double
Dim cCha    As Double
Dim cTen    As Double
Dim i       As Integer

'        If Me.OptCard Then
'            Me.txtMisc = Format((CDbl(Me.txtMarkPercent) / 100) * CDbl(Me.txtSubtotal), "###,##0.00")
'        Else
'            Me.txtMisc = "0.00"
'        End If


            Me.txtTotCash = Format(SumMyList(4), "###,##0.00")
            Me.txtTotCard = Format(SumMyList(5), "###,##0.00")
            Me.txtTotCheck = Format(SumMyList(6), "###,##0.00")
            Me.txtTotOthers = Format(SumMyList(7), "###,##0.00")


            Me.txtTotal = Format(CDbl(Me.txtMisc) + CDbl(Me.txtSubtotal), "###,##0.00")
            Me.txtAmountTendered = Format(TotalPayments, "###,##0.00")
            
            cTot = CDbl(Me.txtTotal)
            cTen = CDbl(Me.txtAmountTendered)
            cBal = cTot - cTen
            If cBal >= 0 Then
                Me.txtAmountBalance = Format(cBal, "###,##0.00")
            Else
                Me.txtAmountBalance = "0.00"
            End If
     
        
End Sub



Function DisplayBalance(Param) As Double
Dim RsStatements As ADODB.Recordset

SQL = "SELECT * FROM tbl_StaInternational WHERE [StatementNo]='" & Param & "'"
Set RsStatements = New ADODB.Recordset
With RsStatements
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               DisplayBalance = .Fields("Balance").Value
            Else
               DisplayBalance = 0
            End If
            .Close
End With
End Function


Function CheckIfCardActivated(Param) As Boolean
Dim RsCheck As ADODB.Recordset
Select Case Param
 Case "Domestic"
    SQL = "SELECT * FROM tbl_Statement WHERE [Snumber]='" & Me.txtStatementNo & "'"
 Case "International"
    SQL = "SELECT * FROM tbl_StaInternational WHERE [StatementNo]='" & Me.txtStatementNo & "'"
 Case "Documents"
    SQL = "SELECT * FROM tbl_PPT WHERE [StatementNo]='" & Me.txtStatementNo & "'"
End Select

Set RsCheck = New ADODB.Recordset
With RsCheck
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            If .Fields("Credit Card Activated").Value = True Then
            CheckIfCardActivated = True
            Else
            CheckIfCardActivated = False
            End If
        End If
       .Close
End With
End Function

Function CheckIfPaid(Param) As Boolean
Dim RsCheck As ADODB.Recordset
Select Case Param
 Case "Domestic"
    SQL = "SELECT * FROM tbl_Statement WHERE [Snumber]='" & Me.txtStatementNo & "'"
 Case "International"
    SQL = "SELECT * FROM tbl_Statement_INTL WHERE [SANo]='" & Me.txtStatementNo & "'"
 Case "Documents"
    SQL = "SELECT * FROM tbl_PPT WHERE [StatementNo]='" & Me.txtStatementNo & "'"
End Select

Set RsCheck = New ADODB.Recordset
With RsCheck
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            If .Fields("Paid").Value = True Then
            CheckIfPaid = True
            Else
            CheckIfPaid = False
            End If
        End If
       .Close
End With

End Function


'SD-1-000000001
Private Sub txtMarkPercent_Change()
Call ReCompute
End Sub

Function Get_OR_LastNum() As String
Dim RsFnumber As ADODB.Recordset
Dim SQL As String

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT * from tbl_OR ORDER BY [ORno] ASC"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               Get_OR_LastNum = RsFnumber("ORno").Value
        Else
               Get_OR_LastNum = "0000000000"
        End If
        .Close
      Set RsFnumber = Nothing
End With

End Function

Private Sub txtReceivedFrom_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub
