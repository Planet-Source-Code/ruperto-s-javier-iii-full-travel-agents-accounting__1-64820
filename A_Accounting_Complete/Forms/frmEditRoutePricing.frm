VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmEditRoutePricing 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Route Pricing"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   Icon            =   "frmEditRoutePricing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   7080
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
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":0CE6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":1B38
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":298A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":3264
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":3B3E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":4418
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":4DE2
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":56BC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":59D6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":62B0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":6B8A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":7464
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":777E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":8058
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":8932
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":920C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":9AE6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":A3C0
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":AC9A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":B574
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":BE4E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":C728
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":D002
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":D8DC
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":E1B6
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":EA90
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":F36A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":FC44
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":1051E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":10DD4
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":116AE
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":11B00
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":11F52
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRoutePricing.frx":14704
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Selection"
      Height          =   7395
      Left            =   75
      TabIndex        =   7
      Top             =   105
      Width           =   6930
      Begin VB.ComboBox txtTicketType 
         Height          =   315
         Left            =   2130
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   795
         Width           =   4605
      End
      Begin VB.TextBox txtShipLines 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   345
         Width           =   4590
      End
      Begin VB.TextBox txtRoute 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1290
         Width           =   4590
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   4530
         Left            =   180
         ScaleHeight     =   4470
         ScaleWidth      =   6495
         TabIndex        =   9
         Top             =   1755
         Width           =   6555
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Important Fill this for Refund"
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   3345
            TabIndex        =   34
            Top             =   105
            Width           =   3015
            Begin VB.TextBox txtNoShowFee 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   40
               Text            =   "0.00"
               Top             =   1455
               Width           =   1470
            End
            Begin VB.TextBox txtVoidFee 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   39
               Text            =   "0.00"
               Top             =   1110
               Width           =   1470
            End
            Begin VB.TextBox txtServiceFee 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   36
               Text            =   "0.00"
               Top             =   360
               Width           =   1470
            End
            Begin VB.TextBox txtRefund 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1380
               TabIndex        =   35
               Text            =   "0.00"
               Top             =   735
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
               TabIndex        =   42
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
               TabIndex        =   41
               Top             =   1095
               Width           =   1260
            End
            Begin VB.Label Label2 
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
               TabIndex        =   38
               Top             =   345
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
               TabIndex        =   37
               Top             =   705
               Width           =   1260
            End
         End
         Begin VB.TextBox txtMisc 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1530
            TabIndex        =   32
            Text            =   "0.00"
            Top             =   2460
            Width           =   1455
         End
         Begin VB.TextBox txtEvat 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0"
            Top             =   3435
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox txtVAT 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2100
            TabIndex        =   25
            Text            =   "0"
            Top             =   2955
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox txtMeals 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1530
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   2100
            Width           =   1470
         End
         Begin VB.TextBox txtGrossFare 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1530
            TabIndex        =   0
            Text            =   "0.00"
            Top             =   195
            Width           =   1470
         End
         Begin VB.TextBox txtInsurance 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1530
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   945
            Width           =   1470
         End
         Begin VB.TextBox txtCommision 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2085
            TabIndex        =   1
            Text            =   "0"
            Top             =   570
            Width           =   915
         End
         Begin VB.TextBox txtASF 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1530
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   1335
            Width           =   1470
         End
         Begin VB.TextBox txtNetFare 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   4005
            Width           =   2490
         End
         Begin VB.TextBox txtTerminalFee 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1530
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   1710
            Width           =   1470
         End
         Begin VB.Label Label1 
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
            Left            =   120
            TabIndex        =   33
            Top             =   2520
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
            Left            =   1770
            TabIndex        =   31
            Top             =   3465
            Visible         =   0   'False
            Width           =   270
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
            Left            =   60
            TabIndex        =   30
            Top             =   3465
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "VAT %  :"
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
            Left            =   60
            TabIndex        =   27
            Top             =   3060
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label18 
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
            Left            =   1770
            TabIndex        =   26
            Top             =   2970
            Visible         =   0   'False
            Width           =   270
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
            Left            =   120
            TabIndex        =   24
            Top             =   2115
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
            Left            =   105
            TabIndex        =   17
            Top             =   240
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
            Left            =   120
            TabIndex        =   16
            Top             =   975
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
            Left            =   135
            TabIndex        =   15
            Top             =   615
            Width           =   1890
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
            Left            =   135
            TabIndex        =   14
            Top             =   1410
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
            Left            =   1515
            TabIndex        =   13
            Top             =   4080
            Width           =   1890
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
            Left            =   120
            TabIndex        =   12
            Top             =   1755
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
            Left            =   1785
            TabIndex        =   11
            Top             =   450
            Width           =   270
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   120
         TabIndex        =   8
         Top             =   6405
         Width           =   6600
         Begin LVbuttons.LaVolpeButton cmdAddSave 
            Height          =   480
            Left            =   105
            TabIndex        =   5
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Save"
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
            MICON           =   "frmEditRoutePricing.frx":15986
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
            Left            =   5010
            TabIndex        =   6
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
            MICON           =   "frmEditRoutePricing.frx":159A2
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
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping / Airline Name :"
         Height          =   285
         Left            =   270
         TabIndex        =   22
         Top             =   390
         Width           =   1890
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Type"
         Height          =   285
         Left            =   300
         TabIndex        =   21
         Top             =   885
         Width           =   1890
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         Height          =   285
         Left            =   315
         TabIndex        =   20
         Top             =   1380
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmEditRoutePricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim RstFareBasis As New ADODB.Recordset
Dim SQL As String


Private Sub cmdAddSave_Click()
'On Error GoTo ErrExit
Dim ask As Integer
ask = MsgBox("Are you sure you want to update this record?", vbInformation + vbYesNo, "ELS TRAVEL & Tours")
If ask = vbYes Then
        With Rs
        cn.BeginTrans
            .Fields(3).Value = FindTicketTypeID(Me.txtTicketType)
            .Fields(4).Value = CDbl(Me.txtGrossFare)
            .Fields(5).Value = CDbl(Me.txtInsurance)
            .Fields(6).Value = CDbl(Me.txtCommision)
            .Fields(7).Value = CDbl(Me.txtASF)
            .Fields(8).Value = CDbl(Me.txtTerminalFee)
            .Fields(9).Value = CDbl(Me.txtNetFare)
            .Fields(10).Value = CDbl(Me.txtMeals)
            .Fields(11).Value = CDbl(Me.txtVat)
            .Fields(12).Value = CDbl(Me.txtEVAT)
            .Fields("Misc").Value = CDbl(Me.txtMisc)
            .Fields("Service Fee").Value = CDbl(Me.txtServiceFee)
            .Fields("Refund Fee").Value = CDbl(Me.txtRefund)
            .Fields("Void Fee").Value = CDbl(Me.txtVoidFee)
            .Fields("Noshow Fee").Value = CDbl(Me.txtNoShowFee)
            .Update
            .Requery
        cn.CommitTrans
        End With
        frmDisplayRoutePricing.Refresh_Grid
        MsgBox "Record Save...", vbInformation
End If
Exit Sub
ErrExit:
cn.RollbackTrans
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Sub ReCalc()
Dim Tmp As Double
Dim TmpComm As Double
Dim TmpCommWithVAT As Double
Dim TmpVAT As Double
Dim TmpEvat As Double

With Me
    TmpEvat = (CDbl(Me.txtGrossFare) + CDbl(Me.txtASF) + CDbl(Me.txtInsurance) + CDbl(Me.txtTerminalFee)) * CDbl(Me.txtEVAT) / 100
    TmpComm = CDbl(Me.txtGrossFare) * (CDbl(Me.txtCommision) / 100)
    
    TmpCommWithVAT = TmpComm * (CDbl(Me.txtVat) / 100)
    
    TmpVAT = (CDbl(Me.txtGrossFare) + CDbl(txtInsurance) + CDbl(txtASF) + CDbl(txtTerminalFee)) * (CDbl(Me.txtVat) / 100)
    
    Tmp = CDbl(Me.txtGrossFare) + CDbl(txtInsurance) + CDbl(Me.txtMeals) + CDbl(txtASF) _
        + CDbl(txtTerminalFee) + CDbl(TmpCommWithVAT) + CDbl(Me.txtMisc) + CDbl(Me.txtEVAT)
        
    .txtNetFare = Format(Tmp, "###,##0.00")
    
End With
End Sub


Private Sub Form_Activate()
Set Rs = New ADODB.Recordset


SQL = "SELECT * FROM tbl_RoutePricing WHERE [RoutePricingID]=" & CDbl(Me.Tag)
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
End With

Me.txtRoute = frmDisplayRoutePricing.DataGrid1.Columns(4).Text & " - " & frmDisplayRoutePricing.DataGrid1.Columns(5).Text
Me.txtShipLines = frmDisplayRoutePricing.DataGrid1.Columns(7).Text

With Me
        .txtGrossFare = IIf(IsNull(Rs.Fields(4)), "0.00", Format(Rs.Fields(4), "###,##0.00"))
        .txtInsurance = IIf(IsNull(Rs.Fields(5)), "0.00", Format(Rs.Fields(5), "###,##0.00"))
        .txtCommision = IIf(IsNull(Rs.Fields(6)), "0.00", Format(Rs.Fields(6), "###,##0.00"))
        .txtASF = IIf(IsNull(Rs.Fields(7)), "0.00", Format(Rs.Fields(7), "###,##0.00"))
        .txtTerminalFee = IIf(IsNull(Rs.Fields(8)), "0.00", Format(Rs.Fields(8), "###,##0.00"))
        .txtNetFare = IIf(IsNull(Rs.Fields(9)), "0.00", Format(Rs.Fields(9), "###,##0.00"))
        .txtMeals = IIf(IsNull(Rs.Fields(10)), "0.00", Format(Rs.Fields(10), "###,##0.00"))
        .txtVat = IIf(IsNull(Rs.Fields(11)), "0.00", Format(Rs.Fields(11), "###,##0.00"))
        .txtEVAT = IIf(IsNull(Rs.Fields(12)), "0.00", Format(Rs.Fields(12), "###,##0.00"))
        .txtMisc = IIf(IsNull(Rs.Fields("Misc")), "0.00", Format(Rs.Fields("Misc"), "###,##0.00"))
        
        .txtServiceFee = IIf(IsNull(Rs.Fields("Service Fee")), "0.00", Format(Rs.Fields("Service Fee"), "###,##0.00"))
        .txtRefund = IIf(IsNull(Rs.Fields("Refund Fee")), "0.00", Format(Rs.Fields("Refund Fee"), "###,##0.00"))
        .txtVoidFee = IIf(IsNull(Rs.Fields("Void Fee")), "0.00", Format(Rs.Fields("Void Fee"), "###,##0.00"))
        .txtNoShowFee = IIf(IsNull(Rs.Fields("Noshow Fee")), "0.00", Format(Rs.Fields("Noshow Fee"), "###,##0.00"))
        
End With
Call FillTicketType
Me.txtTicketType = frmDisplayRoutePricing.DataGrid1.Columns(9).Text
End Sub

Private Sub txtASF_Change()
If IsNumeric(Me.txtASF) Then
ReCalc
End If
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

Private Sub txtGrossFare_Change()
ReCalc
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
If IsNumeric(Me.txtInsurance) Then
ReCalc
End If
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

Private Sub txtMeals_Change()
If IsNumeric(Me.txtMeals) Then
    ReCalc
End If
End Sub

Private Sub txtMeals_GotFocus()
Me.txtMeals = Format(Me.txtMeals, "###,##0.00")
End Sub

Private Sub txtMeals_LostFocus()
Me.txtMeals = Format(Me.txtMeals, "###,##0.00")
End Sub


Private Sub txtMisc_Change()
If IsNumeric(Me.txtMisc) Then
ReCalc
End If

End Sub

Private Sub txtTerminalFee_Change()
If IsNumeric(Me.txtTerminalFee) Then
ReCalc
End If
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

Sub FillTicketType()
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_TicketType"

With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
       If .RecordCount > 0 Then
            Me.txtTicketType.Clear
            Do While Not .EOF
                Me.txtTicketType.AddItem .Fields(1).Value
                .MoveNext
            Loop
       End If
     .Close
     
End With

Set Tmp = Nothing
End Sub

Function FindTicketTypeID(param) As Long
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_TicketType WHERE [Ticket Type]='" & param & "'"
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
       If .RecordCount > 0 Then
            FindTicketTypeID = .Fields(0).Value
       Else
            FindTicketTypeID = -1
       End If
End With
End Function

