VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmStatement 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   5775
      Top             =   90
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
            Picture         =   "Form1.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":006A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":00C8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F1A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17F4
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":20CE
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29A8
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3372
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C4C
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F66
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4840
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":511A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":59F4
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D0E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":65E8
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6EC2
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":779C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8076
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8950
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":922A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9B04
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A3DE
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ACB8
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B592
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BE6C
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C746
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D020
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D8FA
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E1D4
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EAAE
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F364
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FC3E
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10090
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":104E2
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12C94
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdPost 
      Height          =   480
      Left            =   2760
      TabIndex        =   116
      Top             =   540
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
      MICON           =   "Form1.frx":13F16
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
   Begin VB.TextBox txtMisc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13215
      TabIndex        =   579
      Text            =   "0.00"
      Top             =   1770
      Width           =   1620
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   525
      Left            =   13395
      TabIndex        =   115
      Top             =   9345
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   109
      Top             =   1095
      Width           =   15135
      Begin VB.ComboBox CboAccountName 
         Height          =   315
         Left            =   1155
         TabIndex        =   492
         Top             =   615
         Width           =   3225
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Shipping Line Meals Option"
         Height          =   390
         Left            =   9315
         TabIndex        =   491
         Top             =   1605
         Width           =   5745
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   945
         Left            =   4545
         ScaleHeight     =   885
         ScaleWidth      =   4605
         TabIndex        =   256
         Top             =   1050
         Width           =   4665
         Begin VB.TextBox txtCommPercent 
            Alignment       =   2  'Center
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
            Height          =   450
            Left            =   780
            TabIndex        =   257
            Text            =   "0"
            Top             =   330
            Width           =   1095
         End
         Begin LVbuttons.LaVolpeButton cmdSet 
            Height          =   495
            Left            =   2460
            TabIndex        =   259
            Top             =   315
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Manual Set"
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
            MICON           =   "Form1.frx":13F32
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
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   4
            Left            =   1965
            TabIndex        =   260
            Top             =   405
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Commission Percentage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   825
            TabIndex        =   258
            Top             =   0
            Width           =   3105
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   1470
         Left            =   9315
         ScaleHeight     =   1410
         ScaleWidth      =   5670
         TabIndex        =   248
         Top             =   105
         Width           =   5730
         Begin VB.TextBox txtTotComm 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1590
            TabIndex        =   253
            Text            =   "0.00"
            Top             =   90
            Width           =   1380
         End
         Begin VB.TextBox txtTotDiscount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1590
            TabIndex        =   252
            Text            =   "0.00"
            Top             =   525
            Width           =   1380
         End
         Begin VB.TextBox txtTotFare 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   1590
            TabIndex        =   250
            Text            =   "0.00"
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Misc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   578
            Top             =   105
            Width           =   675
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Comm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            TabIndex        =   254
            Top             =   90
            Width           =   2460
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Disc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            TabIndex        =   251
            Top             =   480
            Width           =   2040
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Fare "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   45
            TabIndex        =   249
            Top             =   915
            Width           =   2040
         End
      End
      Begin VB.TextBox Text2 
         Height          =   390
         Left            =   5970
         TabIndex        =   3
         Top             =   600
         Width           =   3210
      End
      Begin VB.TextBox txtAgencyName 
         Height          =   390
         Left            =   5970
         TabIndex        =   1
         Top             =   165
         Width           =   3210
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1455
         Width           =   3240
      End
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1155
         TabIndex        =   2
         Top             =   1020
         Width           =   3210
      End
      Begin VB.TextBox txtNo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1155
         TabIndex        =   111
         Top             =   165
         Width           =   3210
      End
      Begin VB.Label Label2 
         Caption         =   "Tel. No"
         Height          =   390
         Index           =   2
         Left            =   4530
         TabIndex        =   247
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Agency / Name"
         Height          =   390
         Index           =   1
         Left            =   4575
         TabIndex        =   246
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Ship / Airline :"
         Height          =   330
         Left            =   90
         TabIndex        =   114
         Top             =   1485
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Date :"
         Height          =   390
         Left            =   150
         TabIndex        =   113
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Acc. Name:"
         Height          =   390
         Index           =   0
         Left            =   165
         TabIndex        =   112
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "No. :"
         Height          =   390
         Left            =   165
         TabIndex        =   110
         Top             =   180
         Width           =   525
      End
   End
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   13530
      TabIndex        =   117
      Top             =   540
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
      MICON           =   "Form1.frx":13F4E
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
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   480
      Left            =   30
      TabIndex        =   0
      Top             =   540
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
      MICON           =   "Form1.frx":13F6A
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
      Left            =   1395
      TabIndex        =   118
      Top             =   540
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
      MICON           =   "Form1.frx":13F86
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
   Begin LVbuttons.LaVolpeButton cmdRecalc 
      Height          =   480
      Left            =   4125
      TabIndex        =   255
      Top             =   540
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
      MICON           =   "Form1.frx":13FA2
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
   Begin Accounting.ucGradContainer ucGradContainer1 
      Height          =   1155
      Left            =   15
      TabIndex        =   580
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   2037
      IconSize        =   0
      HeaderColor2    =   15523027
      HeaderColor1    =   14334632
      BackColor2      =   16315119
      BackColor1      =   15853021
      BorderColor     =   14070944
      CaptionColor    =   4925975
      Caption         =   "Statement of Account"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderIcon      =   "Form1.frx":13FBE
      Theme           =   1
      Movable         =   -1  'True
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   5910
      Left            =   -15
      TabIndex        =   119
      Top             =   3120
      Width           =   15135
      Begin VB.TextBox txtIssuedBy 
         Enabled         =   0   'False
         Height          =   465
         Left            =   1320
         TabIndex        =   656
         Top             =   5385
         Width           =   3465
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   3315
         TabIndex        =   105
         Text            =   "Combo2"
         Top             =   4980
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   3315
         TabIndex        =   99
         Text            =   "Combo2"
         Top             =   4650
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   3315
         TabIndex        =   93
         Text            =   "Combo2"
         Top             =   4320
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   3315
         TabIndex        =   87
         Text            =   "Combo2"
         Top             =   3990
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   3315
         TabIndex        =   81
         Text            =   "Combo2"
         Top             =   3660
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   3315
         TabIndex        =   75
         Text            =   "Combo2"
         Top             =   3330
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   3315
         TabIndex        =   69
         Text            =   "Combo2"
         Top             =   3000
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   3315
         TabIndex        =   64
         Text            =   "Combo2"
         Top             =   2670
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3315
         TabIndex        =   61
         Text            =   "Combo2"
         Top             =   2340
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   3315
         TabIndex        =   58
         Text            =   "Combo2"
         Top             =   2010
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3315
         TabIndex        =   55
         Text            =   "Combo2"
         Top             =   1680
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3315
         TabIndex        =   52
         Text            =   "Combo2"
         Top             =   1350
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3315
         TabIndex        =   29
         Text            =   "REGULAR"
         Top             =   1020
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3315
         TabIndex        =   18
         Text            =   "ECONOMY"
         Top             =   690
         Width           =   750
      End
      Begin VB.ComboBox cboTicketType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3315
         TabIndex        =   7
         Text            =   "JETSETTER"
         Top             =   375
         Width           =   750
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   275
         Text            =   "0"
         Top             =   375
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   274
         Text            =   "1234567890"
         Top             =   705
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   273
         Text            =   "1234567890"
         Top             =   1035
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   272
         Text            =   "1234567890"
         Top             =   1365
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   271
         Text            =   "1234567890"
         Top             =   1695
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   270
         Text            =   "1234567890"
         Top             =   2025
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   269
         Text            =   "1234567890"
         Top             =   2355
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   268
         Text            =   "1234567890"
         Top             =   2685
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   267
         Text            =   "1234567890"
         Top             =   3015
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   266
         Text            =   "1234567890"
         Top             =   3345
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   265
         Text            =   "1234567890"
         Top             =   3675
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   264
         Text            =   "1234567890"
         Top             =   4005
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   263
         Text            =   "1234567890"
         Top             =   4335
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   262
         Text            =   "1234567890"
         Top             =   4665
         Width           =   795
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   12900
         MaxLength       =   10
         TabIndex        =   261
         Text            =   "1234567890"
         Top             =   4995
         Width           =   795
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   244
         Text            =   "0"
         Top             =   4995
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   243
         Text            =   "0"
         Top             =   4665
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   242
         Text            =   "0"
         Top             =   4335
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   241
         Text            =   "0"
         Top             =   4005
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   240
         Text            =   "0"
         Top             =   3675
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   239
         Text            =   "0"
         Top             =   3345
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   238
         Text            =   "0"
         Top             =   3015
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   237
         Text            =   "0"
         Top             =   2685
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   236
         Text            =   "0"
         Top             =   2355
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   235
         Text            =   "0"
         Top             =   2025
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   234
         Text            =   "0"
         Top             =   1695
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   233
         Text            =   "0"
         Top             =   1365
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   232
         Text            =   "0"
         Top             =   1035
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   231
         Text            =   "0"
         Top             =   705
         Width           =   630
      End
      Begin VB.TextBox txtComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   12285
         MaxLength       =   4
         TabIndex        =   230
         Text            =   "0"
         Top             =   375
         Width           =   630
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   229
         Text            =   "1234567890"
         Top             =   375
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   228
         Text            =   "1234567890"
         Top             =   705
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   227
         Text            =   "1234567890"
         Top             =   1035
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   226
         Text            =   "1234567890"
         Top             =   1365
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   225
         Text            =   "1234567890"
         Top             =   1695
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   224
         Text            =   "1234567890"
         Top             =   2025
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   223
         Text            =   "1234567890"
         Top             =   2355
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   222
         Text            =   "1234567890"
         Top             =   2685
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   221
         Text            =   "1234567890"
         Top             =   3015
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   220
         Text            =   "1234567890"
         Top             =   3345
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   219
         Text            =   "1234567890"
         Top             =   3675
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   218
         Text            =   "1234567890"
         Top             =   4005
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   217
         Text            =   "1234567890"
         Top             =   4335
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   216
         Text            =   "1234567890"
         Top             =   4665
         Width           =   795
      End
      Begin VB.TextBox txtGrossFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   11505
         MaxLength       =   10
         TabIndex        =   215
         Text            =   "1234567890"
         Top             =   4995
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   212
         Text            =   "1234567890"
         Top             =   4995
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   211
         Text            =   "1234567890"
         Top             =   4665
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   210
         Text            =   "1234567890"
         Top             =   4335
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   209
         Text            =   "1234567890"
         Top             =   4005
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   208
         Text            =   "1234567890"
         Top             =   3675
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   207
         Text            =   "1234567890"
         Top             =   3345
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   206
         Text            =   "1234567890"
         Top             =   3015
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   205
         Text            =   "1234567890"
         Top             =   2685
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   204
         Text            =   "1234567890"
         Top             =   2355
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   203
         Text            =   "1234567890"
         Top             =   2025
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   202
         Text            =   "1234567890"
         Top             =   1695
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   201
         Text            =   "1234567890"
         Top             =   1365
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   200
         Text            =   "1234567890"
         Top             =   1035
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   199
         Text            =   "1234567890"
         Top             =   705
         Width           =   795
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   14145
         MaxLength       =   10
         TabIndex        =   198
         Text            =   "1234567890"
         Top             =   375
         Width           =   795
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   197
         Text            =   "1234"
         Top             =   4995
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   196
         Text            =   "1234"
         Top             =   4665
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   195
         Text            =   "1234"
         Top             =   4335
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   194
         Text            =   "1234"
         Top             =   4005
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   193
         Text            =   "1234"
         Top             =   3675
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   192
         Text            =   "1234"
         Top             =   3345
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   191
         Text            =   "1234"
         Top             =   3015
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   190
         Text            =   "1234"
         Top             =   2685
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   189
         Text            =   "1234"
         Top             =   2355
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   188
         Text            =   "1234"
         Top             =   2025
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   187
         Text            =   "1234"
         Top             =   1695
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   186
         Text            =   "1234"
         Top             =   1365
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   185
         Text            =   "1234"
         Top             =   1035
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   184
         Text            =   "900"
         Top             =   705
         Width           =   480
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   13680
         MaxLength       =   4
         TabIndex        =   183
         Text            =   "900"
         Top             =   375
         Width           =   480
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   166
         Text            =   "99/99/9999"
         Top             =   4995
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   165
         Text            =   "99/99/9999"
         Top             =   4665
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   164
         Text            =   "99/99/9999"
         Top             =   4335
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   163
         Text            =   "99/99/9999"
         Top             =   3990
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   162
         Text            =   "99/99/9999"
         Top             =   3660
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   161
         Text            =   "99/99/9999"
         Top             =   3330
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   160
         Text            =   "99/99/9999"
         Top             =   3000
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   159
         Text            =   "99/99/9999"
         Top             =   2670
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   158
         Text            =   "99/99/9999"
         Top             =   2340
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   157
         Text            =   "99/99/9999"
         Top             =   2010
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   156
         Text            =   "99/99/9999"
         Top             =   1680
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   155
         Text            =   "99/99/9999"
         Top             =   1350
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   154
         Text            =   "99/99/9999"
         Top             =   1020
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   25
         Text            =   "99/99/9999"
         Top             =   690
         Width           =   870
      End
      Begin VB.TextBox txtDepDate3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   9660
         MaxLength       =   17
         TabIndex        =   14
         Text            =   "99/99/9999"
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   153
         Text            =   "99/99/9999"
         Top             =   4995
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   152
         Text            =   "99/99/9999"
         Top             =   4665
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   151
         Text            =   "99/99/9999"
         Top             =   4335
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   150
         Text            =   "99/99/9999"
         Top             =   3990
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   149
         Text            =   "99/99/9999"
         Top             =   3660
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   148
         Text            =   "99/99/9999"
         Top             =   3330
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   147
         Text            =   "99/99/9999"
         Top             =   3000
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   146
         Text            =   "99/99/9999"
         Top             =   2670
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   145
         Text            =   "99/99/9999"
         Top             =   2340
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   144
         Text            =   "99/99/9999"
         Top             =   2010
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   143
         Text            =   "99/99/9999"
         Top             =   1680
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   142
         Text            =   "99/99/9999"
         Top             =   1350
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   141
         Text            =   "99/99/9999"
         Top             =   1020
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   24
         Text            =   "99/99/9999"
         Top             =   690
         Width           =   870
      End
      Begin VB.TextBox txtDepDate2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   8820
         MaxLength       =   17
         TabIndex        =   13
         Text            =   "99/99/9999"
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   140
         Text            =   "99/99/9999"
         Top             =   4995
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   139
         Text            =   "99/99/9999"
         Top             =   4665
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   138
         Text            =   "99/99/9999"
         Top             =   4335
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   137
         Text            =   "99/99/9999"
         Top             =   3990
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   136
         Text            =   "99/99/9999"
         Top             =   3660
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   135
         Text            =   "99/99/9999"
         Top             =   3330
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   134
         Text            =   "99/99/9999"
         Top             =   3000
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   133
         Text            =   "99/99/9999"
         Top             =   2670
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   132
         Text            =   "99/99/9999"
         Top             =   2340
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   131
         Text            =   "99/99/9999"
         Top             =   2010
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   130
         Text            =   "99/99/9999"
         Top             =   1680
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   129
         Text            =   "99/99/9999"
         Top             =   1350
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   128
         Text            =   "99/99/9999"
         Top             =   1020
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   23
         Text            =   "99/99/9999"
         Top             =   690
         Width           =   870
      End
      Begin VB.TextBox txtDepDate1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   7980
         MaxLength       =   17
         TabIndex        =   12
         Text            =   "99/99/9999 10:50A"
         Top             =   390
         Width           =   870
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   5970
         TabIndex        =   108
         Text            =   "Combo2"
         Top             =   4980
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   5970
         TabIndex        =   102
         Text            =   "Combo2"
         Top             =   4650
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   5970
         TabIndex        =   96
         Text            =   "Combo2"
         Top             =   4320
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   5970
         TabIndex        =   90
         Text            =   "Combo2"
         Top             =   3990
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   5970
         TabIndex        =   84
         Text            =   "Combo2"
         Top             =   3660
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   5970
         TabIndex        =   78
         Text            =   "Combo2"
         Top             =   3330
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   5970
         TabIndex        =   72
         Text            =   "Combo2"
         Top             =   3000
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   5970
         TabIndex        =   66
         Text            =   "Combo2"
         Top             =   2670
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   5970
         TabIndex        =   48
         Text            =   "Combo2"
         Top             =   2340
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   5970
         TabIndex        =   44
         Text            =   "Combo2"
         Top             =   2010
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   5970
         TabIndex        =   40
         Text            =   "Combo2"
         Top             =   1680
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   5970
         TabIndex        =   36
         Text            =   "Combo2"
         Top             =   1350
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   5970
         TabIndex        =   32
         Text            =   "Combo2"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   5970
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   690
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   4995
         TabIndex        =   107
         Text            =   "Combo2"
         Top             =   4980
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   4995
         TabIndex        =   101
         Text            =   "Combo2"
         Top             =   4650
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   4995
         TabIndex        =   95
         Text            =   "Combo2"
         Top             =   4320
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   4995
         TabIndex        =   89
         Text            =   "Combo2"
         Top             =   3990
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   4995
         TabIndex        =   83
         Text            =   "Combo2"
         Top             =   3660
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   4995
         TabIndex        =   77
         Text            =   "Combo2"
         Top             =   3330
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   4995
         TabIndex        =   71
         Text            =   "Combo2"
         Top             =   3000
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   4995
         TabIndex        =   125
         Text            =   "Combo2"
         Top             =   2670
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   4995
         TabIndex        =   47
         Text            =   "Combo2"
         Top             =   2340
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   4995
         TabIndex        =   43
         Text            =   "Combo2"
         Top             =   2010
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   4995
         TabIndex        =   39
         Text            =   "Combo2"
         Top             =   1680
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   4995
         TabIndex        =   35
         Text            =   "Combo2"
         Top             =   1350
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   4995
         TabIndex        =   31
         Text            =   "Combo2"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   4995
         TabIndex        =   20
         Text            =   "Combo2"
         Top             =   690
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   4020
         TabIndex        =   106
         Text            =   "Combo2"
         Top             =   4980
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   4020
         TabIndex        =   100
         Text            =   "Combo2"
         Top             =   4650
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   4020
         TabIndex        =   94
         Text            =   "Combo2"
         Top             =   4320
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   4020
         TabIndex        =   88
         Text            =   "Combo2"
         Top             =   3990
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   4020
         TabIndex        =   82
         Text            =   "Combo2"
         Top             =   3660
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   4020
         TabIndex        =   76
         Text            =   "Combo2"
         Top             =   3330
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   4020
         TabIndex        =   70
         Text            =   "Combo2"
         Top             =   3000
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   4020
         TabIndex        =   65
         Text            =   "Combo2"
         Top             =   2670
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   4020
         TabIndex        =   46
         Text            =   "Combo2"
         Top             =   2340
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   4020
         TabIndex        =   42
         Text            =   "Combo2"
         Top             =   2010
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   4020
         TabIndex        =   38
         Text            =   "Combo2"
         Top             =   1680
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   4020
         TabIndex        =   34
         Text            =   "Combo2"
         Top             =   1350
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   4020
         TabIndex        =   30
         Text            =   "Combo2"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   4020
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   690
         Width           =   1065
      End
      Begin VB.ComboBox cboDest2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   5970
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   375
         Width           =   1065
      End
      Begin VB.ComboBox cboDest1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   4995
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   375
         Width           =   1065
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   4020
         TabIndex        =   8
         Text            =   "cboFrom"
         Top             =   375
         Width           =   1065
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   104
         Top             =   4995
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   98
         Top             =   4665
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   92
         Top             =   4335
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   86
         Top             =   4005
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   80
         Top             =   3675
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   74
         Top             =   3345
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   68
         Top             =   3015
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   63
         Top             =   2685
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   60
         Top             =   2355
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   57
         Top             =   2025
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   54
         Top             =   1695
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   51
         Top             =   1365
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   28
         Top             =   1035
         Width           =   2280
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   17
         Top             =   705
         Width           =   2280
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   45
         MaxLength       =   11
         TabIndex        =   103
         Top             =   4995
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   45
         MaxLength       =   11
         TabIndex        =   97
         Top             =   4665
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   45
         MaxLength       =   11
         TabIndex        =   91
         Top             =   4335
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   45
         MaxLength       =   11
         TabIndex        =   85
         Top             =   4005
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   45
         MaxLength       =   11
         TabIndex        =   79
         Top             =   3675
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   45
         MaxLength       =   11
         TabIndex        =   73
         Top             =   3345
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   45
         MaxLength       =   11
         TabIndex        =   67
         Top             =   3015
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   45
         MaxLength       =   11
         TabIndex        =   62
         Top             =   2685
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   45
         MaxLength       =   11
         TabIndex        =   59
         Top             =   2355
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   45
         MaxLength       =   11
         TabIndex        =   56
         Top             =   2025
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   45
         MaxLength       =   11
         TabIndex        =   53
         Top             =   1695
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   45
         MaxLength       =   11
         TabIndex        =   50
         Top             =   1365
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   45
         MaxLength       =   11
         TabIndex        =   27
         Top             =   1035
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   45
         MaxLength       =   11
         TabIndex        =   16
         Top             =   705
         Width           =   1035
      End
      Begin VB.TextBox txtTicketNo 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   45
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "1234567890"
         Top             =   375
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   10515
         TabIndex        =   170
         Text            =   "Combo2"
         Top             =   4980
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   10515
         TabIndex        =   171
         Text            =   "Combo2"
         Top             =   4650
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   10515
         TabIndex        =   172
         Text            =   "Combo2"
         Top             =   4320
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   10515
         TabIndex        =   173
         Text            =   "Combo2"
         Top             =   3990
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   10515
         TabIndex        =   174
         Text            =   "Combo2"
         Top             =   3660
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   10515
         TabIndex        =   175
         Text            =   "Combo2"
         Top             =   3330
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   10515
         TabIndex        =   176
         Text            =   "Combo2"
         Top             =   3000
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   10515
         TabIndex        =   177
         Text            =   "Combo2"
         Top             =   2670
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   10515
         TabIndex        =   178
         Text            =   "Combo2"
         Top             =   2340
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   10515
         TabIndex        =   179
         Text            =   "Combo2"
         Top             =   2010
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   10515
         TabIndex        =   180
         Text            =   "Combo2"
         Top             =   1680
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   10515
         TabIndex        =   181
         Text            =   "INFANT"
         Top             =   1350
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   10515
         TabIndex        =   182
         Text            =   "FULL"
         Top             =   1020
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   10515
         TabIndex        =   26
         Text            =   "STUDENT"
         Top             =   690
         Width           =   1020
      End
      Begin VB.ComboBox cboPassengerType 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   10515
         TabIndex        =   15
         Text            =   "SENIOR"
         Top             =   375
         Width           =   1020
      End
      Begin VB.TextBox txtPassengerName 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   6
         Top             =   375
         Width           =   2280
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   6975
         TabIndex        =   493
         Text            =   "Combo2"
         Top             =   4980
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   6975
         TabIndex        =   494
         Text            =   "Combo2"
         Top             =   4650
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   6975
         TabIndex        =   495
         Text            =   "Combo2"
         Top             =   4320
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   6975
         TabIndex        =   496
         Text            =   "Combo2"
         Top             =   3990
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   6975
         TabIndex        =   497
         Text            =   "Combo2"
         Top             =   3660
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   6975
         TabIndex        =   498
         Text            =   "Combo2"
         Top             =   3330
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   6975
         TabIndex        =   499
         Text            =   "Combo2"
         Top             =   3000
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   6975
         TabIndex        =   500
         Text            =   "Combo2"
         Top             =   2670
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   6975
         TabIndex        =   49
         Text            =   "Combo2"
         Top             =   2340
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   6975
         TabIndex        =   45
         Text            =   "Combo2"
         Top             =   2010
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   6975
         TabIndex        =   41
         Text            =   "Combo2"
         Top             =   1680
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   6975
         TabIndex        =   37
         Text            =   "Combo2"
         Top             =   1350
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   6975
         TabIndex        =   33
         Text            =   "Combo2"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   6975
         TabIndex        =   22
         Text            =   "Combo2"
         Top             =   690
         Width           =   1065
      End
      Begin VB.ComboBox cboDest3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   6975
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Issued by :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         TabIndex        =   657
         Top             =   5400
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Route 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   15
         Left            =   7005
         TabIndex        =   501
         Top             =   135
         Width           =   990
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "TicketType"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   3330
         TabIndex        =   490
         Top             =   135
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "NET"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   14
         Left            =   12930
         TabIndex        =   276
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "PASSENGER NAME"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1065
         TabIndex        =   121
         Top             =   135
         Width           =   2250
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Comm%"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   13
         Left            =   12315
         TabIndex        =   245
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "FARE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   12
         Left            =   14160
         TabIndex        =   214
         Top             =   135
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "GROSS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   11
         Left            =   11520
         TabIndex        =   213
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Dep Date/Time 3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   9690
         TabIndex        =   169
         Top             =   135
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Dep Date/Time 2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   8850
         TabIndex        =   168
         Top             =   135
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Dep Date/Time 1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   8010
         TabIndex        =   167
         Top             =   135
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Disc %"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   7
         Left            =   13680
         TabIndex        =   127
         Top             =   135
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "VOID"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   10530
         TabIndex        =   126
         Top             =   135
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Route 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   6000
         TabIndex        =   124
         Top             =   135
         Width           =   990
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Route 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   5025
         TabIndex        =   123
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Route 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   4050
         TabIndex        =   122
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "TICKET NO."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   120
         Top             =   135
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   5340
      Left            =   180
      TabIndex        =   277
      Top             =   4530
      Width           =   3165
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   610
         Text            =   "0"
         Top             =   4860
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   609
         Text            =   "0"
         Top             =   4530
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   608
         Text            =   "0"
         Top             =   4200
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   607
         Text            =   "0"
         Top             =   3870
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   606
         Text            =   "0"
         Top             =   3540
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   605
         Text            =   "0"
         Top             =   3210
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   604
         Text            =   "0"
         Top             =   2880
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   603
         Text            =   "0"
         Top             =   2550
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   602
         Text            =   "0"
         Top             =   2220
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   601
         Text            =   "0"
         Top             =   1890
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   600
         Text            =   "0"
         Top             =   1560
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   599
         Text            =   "0"
         Top             =   1230
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   598
         Text            =   "0"
         Top             =   900
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   597
         Text            =   "0"
         Top             =   570
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   596
         Text            =   "0"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   595
         Text            =   "0"
         Top             =   4875
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   594
         Text            =   "0"
         Top             =   4545
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   593
         Text            =   "0"
         Top             =   4215
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   592
         Text            =   "0"
         Top             =   3885
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   591
         Text            =   "0"
         Top             =   3555
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   590
         Text            =   "0"
         Top             =   3225
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   589
         Text            =   "0"
         Top             =   2895
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   588
         Text            =   "0"
         Top             =   2565
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   587
         Text            =   "0"
         Top             =   2235
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   586
         Text            =   "0"
         Top             =   1905
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   585
         Text            =   "0"
         Top             =   1575
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   584
         Text            =   "0"
         Top             =   1245
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   583
         Text            =   "0"
         Top             =   915
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   582
         Text            =   "0"
         Top             =   585
         Width           =   450
      End
      Begin VB.TextBox txtMeals1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   581
         Text            =   "0"
         Top             =   255
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   337
         Text            =   "0"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   336
         Text            =   "0"
         Top             =   570
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   335
         Text            =   "0"
         Top             =   900
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   334
         Text            =   "0"
         Top             =   1230
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   333
         Text            =   "0"
         Top             =   1560
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   332
         Text            =   "0"
         Top             =   1890
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   331
         Text            =   "0"
         Top             =   2220
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   330
         Text            =   "0"
         Top             =   2550
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   329
         Text            =   "0"
         Top             =   2880
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   328
         Text            =   "0"
         Top             =   3210
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   327
         Text            =   "0"
         Top             =   3540
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   326
         Text            =   "0"
         Top             =   3870
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   325
         Text            =   "0"
         Top             =   4200
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   324
         Text            =   "0"
         Top             =   4530
         Width           =   450
      End
      Begin VB.TextBox txtTerminalFee1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   323
         Text            =   "0"
         Top             =   4860
         Width           =   450
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   322
         Text            =   "0"
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   321
         Text            =   "0"
         Top             =   585
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   320
         Text            =   "0"
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   319
         Text            =   "0"
         Top             =   1245
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   318
         Text            =   "0"
         Top             =   1575
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   317
         Text            =   "0"
         Top             =   1905
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   316
         Text            =   "0"
         Top             =   2235
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   315
         Text            =   "0"
         Top             =   2565
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   314
         Text            =   "0"
         Top             =   2895
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   313
         Text            =   "0"
         Top             =   3225
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   312
         Text            =   "0"
         Top             =   3555
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   311
         Text            =   "0"
         Top             =   3885
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   310
         Text            =   "0"
         Top             =   4215
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   309
         Text            =   "0"
         Top             =   4545
         Width           =   495
      End
      Begin VB.TextBox txtASF1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   308
         Text            =   "0"
         Top             =   4875
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   615
         MaxLength       =   4
         TabIndex        =   307
         Text            =   "0"
         Top             =   4875
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   615
         MaxLength       =   4
         TabIndex        =   306
         Text            =   "0"
         Top             =   4545
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   615
         MaxLength       =   4
         TabIndex        =   305
         Text            =   "0"
         Top             =   4215
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   615
         MaxLength       =   4
         TabIndex        =   304
         Text            =   "0"
         Top             =   3885
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   615
         MaxLength       =   4
         TabIndex        =   303
         Text            =   "0"
         Top             =   3555
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   615
         MaxLength       =   4
         TabIndex        =   302
         Text            =   "0"
         Top             =   3225
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   615
         MaxLength       =   4
         TabIndex        =   301
         Text            =   "0"
         Top             =   2895
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   615
         MaxLength       =   4
         TabIndex        =   300
         Text            =   "0"
         Top             =   2565
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   615
         MaxLength       =   4
         TabIndex        =   299
         Text            =   "0"
         Top             =   2235
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   615
         MaxLength       =   4
         TabIndex        =   298
         Text            =   "0"
         Top             =   1905
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   615
         MaxLength       =   4
         TabIndex        =   297
         Text            =   "0"
         Top             =   1575
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   615
         MaxLength       =   4
         TabIndex        =   296
         Text            =   "0"
         Top             =   1245
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   615
         MaxLength       =   4
         TabIndex        =   295
         Text            =   "0"
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   615
         MaxLength       =   4
         TabIndex        =   294
         Text            =   "0"
         Top             =   585
         Width           =   495
      End
      Begin VB.TextBox txtInsurance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   615
         MaxLength       =   4
         TabIndex        =   293
         Text            =   "0"
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   105
         MaxLength       =   4
         TabIndex        =   292
         Text            =   "0"
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   105
         MaxLength       =   4
         TabIndex        =   291
         Text            =   "0"
         Top             =   585
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   105
         MaxLength       =   4
         TabIndex        =   290
         Text            =   "0"
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   105
         MaxLength       =   4
         TabIndex        =   289
         Text            =   "0"
         Top             =   1245
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   105
         MaxLength       =   4
         TabIndex        =   288
         Text            =   "0"
         Top             =   1575
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   105
         MaxLength       =   4
         TabIndex        =   287
         Text            =   "0"
         Top             =   1905
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   105
         MaxLength       =   4
         TabIndex        =   286
         Text            =   "0"
         Top             =   2235
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   105
         MaxLength       =   4
         TabIndex        =   285
         Text            =   "0"
         Top             =   2565
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   105
         MaxLength       =   4
         TabIndex        =   284
         Text            =   "0"
         Top             =   2895
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   105
         MaxLength       =   4
         TabIndex        =   283
         Text            =   "0"
         Top             =   3225
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   105
         MaxLength       =   4
         TabIndex        =   282
         Text            =   "0"
         Top             =   3555
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   105
         MaxLength       =   4
         TabIndex        =   281
         Text            =   "0"
         Top             =   3885
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   105
         MaxLength       =   4
         TabIndex        =   280
         Text            =   "0"
         Top             =   4215
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   105
         MaxLength       =   4
         TabIndex        =   279
         Text            =   "0"
         Top             =   4545
         Width           =   495
      End
      Begin VB.TextBox txtCompaComm1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   105
         MaxLength       =   4
         TabIndex        =   278
         Text            =   "0"
         Top             =   4875
         Width           =   495
      End
   End
   Begin VB.Frame Frame55 
      Caption         =   "Frame3"
      Height          =   5340
      Left            =   10230
      TabIndex        =   502
      Top             =   3825
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   655
         Text            =   "0"
         Top             =   4860
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   654
         Text            =   "0"
         Top             =   4530
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   653
         Text            =   "0"
         Top             =   4200
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   652
         Text            =   "0"
         Top             =   3870
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   651
         Text            =   "0"
         Top             =   3540
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   650
         Text            =   "0"
         Top             =   3210
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   649
         Text            =   "0"
         Top             =   2880
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   648
         Text            =   "0"
         Top             =   2550
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   647
         Text            =   "0"
         Top             =   2220
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   646
         Text            =   "0"
         Top             =   1890
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   645
         Text            =   "0"
         Top             =   1545
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   644
         Text            =   "0"
         Top             =   1230
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   643
         Text            =   "0"
         Top             =   900
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   642
         Text            =   "0"
         Top             =   570
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   641
         Text            =   "0"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   105
         MaxLength       =   4
         TabIndex        =   577
         Text            =   "0"
         Top             =   4875
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   105
         MaxLength       =   4
         TabIndex        =   576
         Text            =   "0"
         Top             =   4545
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   105
         MaxLength       =   4
         TabIndex        =   575
         Text            =   "0"
         Top             =   4215
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   105
         MaxLength       =   4
         TabIndex        =   574
         Text            =   "0"
         Top             =   3885
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   105
         MaxLength       =   4
         TabIndex        =   573
         Text            =   "0"
         Top             =   3555
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   105
         MaxLength       =   4
         TabIndex        =   572
         Text            =   "0"
         Top             =   3225
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   105
         MaxLength       =   4
         TabIndex        =   571
         Text            =   "0"
         Top             =   2895
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   105
         MaxLength       =   4
         TabIndex        =   570
         Text            =   "0"
         Top             =   2565
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   105
         MaxLength       =   4
         TabIndex        =   569
         Text            =   "0"
         Top             =   2235
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   105
         MaxLength       =   4
         TabIndex        =   568
         Text            =   "0"
         Top             =   1905
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   105
         MaxLength       =   4
         TabIndex        =   567
         Text            =   "0"
         Top             =   1575
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   105
         MaxLength       =   4
         TabIndex        =   566
         Text            =   "0"
         Top             =   1245
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   105
         MaxLength       =   4
         TabIndex        =   565
         Text            =   "0"
         Top             =   915
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   105
         MaxLength       =   4
         TabIndex        =   564
         Text            =   "0"
         Top             =   585
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   105
         MaxLength       =   4
         TabIndex        =   563
         Text            =   "0"
         Top             =   255
         Width           =   540
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   645
         MaxLength       =   4
         TabIndex        =   562
         Text            =   "0"
         Top             =   255
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   645
         MaxLength       =   4
         TabIndex        =   561
         Text            =   "0"
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   645
         MaxLength       =   4
         TabIndex        =   560
         Text            =   "0"
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   645
         MaxLength       =   4
         TabIndex        =   559
         Text            =   "0"
         Top             =   1245
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   645
         MaxLength       =   4
         TabIndex        =   558
         Text            =   "0"
         Top             =   1575
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   645
         MaxLength       =   4
         TabIndex        =   557
         Text            =   "0"
         Top             =   1905
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   645
         MaxLength       =   4
         TabIndex        =   556
         Text            =   "0"
         Top             =   2235
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   645
         MaxLength       =   4
         TabIndex        =   555
         Text            =   "0"
         Top             =   2565
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   645
         MaxLength       =   4
         TabIndex        =   554
         Text            =   "0"
         Top             =   2895
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   645
         MaxLength       =   4
         TabIndex        =   553
         Text            =   "0"
         Top             =   3225
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   645
         MaxLength       =   4
         TabIndex        =   552
         Text            =   "0"
         Top             =   3555
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   645
         MaxLength       =   4
         TabIndex        =   551
         Text            =   "0"
         Top             =   3885
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   645
         MaxLength       =   4
         TabIndex        =   550
         Text            =   "0"
         Top             =   4215
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   645
         MaxLength       =   4
         TabIndex        =   549
         Text            =   "0"
         Top             =   4545
         Width           =   555
      End
      Begin VB.TextBox txtInsurance4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   645
         MaxLength       =   4
         TabIndex        =   548
         Text            =   "0"
         Top             =   4875
         Width           =   555
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   547
         Text            =   "0"
         Top             =   4875
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   546
         Text            =   "0"
         Top             =   4545
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   545
         Text            =   "0"
         Top             =   4215
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   544
         Text            =   "0"
         Top             =   3885
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   543
         Text            =   "0"
         Top             =   3555
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   542
         Text            =   "0"
         Top             =   3225
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   541
         Text            =   "0"
         Top             =   2895
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   540
         Text            =   "0"
         Top             =   2565
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   539
         Text            =   "0"
         Top             =   2235
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   538
         Text            =   "0"
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   537
         Text            =   "0"
         Top             =   1575
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   536
         Text            =   "0"
         Top             =   1245
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   535
         Text            =   "0"
         Top             =   915
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   534
         Text            =   "0"
         Top             =   585
         Width           =   525
      End
      Begin VB.TextBox txtASF4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   533
         Text            =   "0"
         Top             =   255
         Width           =   525
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   532
         Text            =   "0"
         Top             =   4860
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   531
         Text            =   "0"
         Top             =   4530
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   530
         Text            =   "0"
         Top             =   4200
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   529
         Text            =   "0"
         Top             =   3870
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   528
         Text            =   "0"
         Top             =   3540
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   527
         Text            =   "0"
         Top             =   3210
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   526
         Text            =   "0"
         Top             =   2880
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   525
         Text            =   "0"
         Top             =   2550
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   524
         Text            =   "0"
         Top             =   2220
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   523
         Text            =   "0"
         Top             =   1890
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   522
         Text            =   "0"
         Top             =   1560
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   521
         Text            =   "0"
         Top             =   1230
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   520
         Text            =   "0"
         Top             =   900
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   519
         Text            =   "0"
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   518
         Text            =   "0"
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   517
         Text            =   "0"
         Top             =   4860
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   516
         Text            =   "0"
         Top             =   4530
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   515
         Text            =   "0"
         Top             =   4200
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   514
         Text            =   "0"
         Top             =   3870
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   513
         Text            =   "0"
         Top             =   3540
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   512
         Text            =   "0"
         Top             =   3210
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   511
         Text            =   "0"
         Top             =   2880
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   510
         Text            =   "0"
         Top             =   2550
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   509
         Text            =   "0"
         Top             =   2220
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   508
         Text            =   "0"
         Top             =   1890
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   507
         Text            =   "0"
         Top             =   1560
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   506
         Text            =   "0"
         Top             =   1230
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   505
         Text            =   "0"
         Top             =   900
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   504
         Text            =   "0"
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox txtMeals4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   503
         Text            =   "0"
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame3"
      Height          =   5340
      Left            =   6870
      TabIndex        =   414
      Top             =   3840
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   640
         Text            =   "0"
         Top             =   4875
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   639
         Text            =   "0"
         Top             =   4545
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   638
         Text            =   "0"
         Top             =   4215
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   637
         Text            =   "0"
         Top             =   3885
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   636
         Text            =   "0"
         Top             =   3555
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   635
         Text            =   "0"
         Top             =   3225
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   634
         Text            =   "0"
         Top             =   2895
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   633
         Text            =   "0"
         Top             =   2565
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   632
         Text            =   "0"
         Top             =   2235
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   631
         Text            =   "0"
         Top             =   1905
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   630
         Text            =   "0"
         Top             =   1575
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   629
         Text            =   "0"
         Top             =   1245
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   628
         Text            =   "0"
         Top             =   915
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   627
         Text            =   "0"
         Top             =   585
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2775
         MaxLength       =   4
         TabIndex        =   626
         Text            =   "0"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   105
         MaxLength       =   4
         TabIndex        =   489
         Text            =   "0"
         Top             =   4875
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   105
         MaxLength       =   4
         TabIndex        =   488
         Text            =   "0"
         Top             =   4545
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   105
         MaxLength       =   4
         TabIndex        =   487
         Text            =   "0"
         Top             =   4215
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   105
         MaxLength       =   4
         TabIndex        =   486
         Text            =   "0"
         Top             =   3885
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   105
         MaxLength       =   4
         TabIndex        =   485
         Text            =   "0"
         Top             =   3555
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   105
         MaxLength       =   4
         TabIndex        =   484
         Text            =   "0"
         Top             =   3225
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   105
         MaxLength       =   4
         TabIndex        =   483
         Text            =   "0"
         Top             =   2895
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   105
         MaxLength       =   4
         TabIndex        =   482
         Text            =   "0"
         Top             =   2565
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   105
         MaxLength       =   4
         TabIndex        =   481
         Text            =   "0"
         Top             =   2235
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   105
         MaxLength       =   4
         TabIndex        =   480
         Text            =   "0"
         Top             =   1905
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   105
         MaxLength       =   4
         TabIndex        =   479
         Text            =   "0"
         Top             =   1575
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   105
         MaxLength       =   4
         TabIndex        =   478
         Text            =   "0"
         Top             =   1245
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   105
         MaxLength       =   4
         TabIndex        =   477
         Text            =   "0"
         Top             =   915
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   105
         MaxLength       =   4
         TabIndex        =   476
         Text            =   "0"
         Top             =   585
         Width           =   540
      End
      Begin VB.TextBox txtCompaComm3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   105
         MaxLength       =   4
         TabIndex        =   475
         Text            =   "0"
         Top             =   255
         Width           =   540
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   645
         MaxLength       =   4
         TabIndex        =   474
         Text            =   "0"
         Top             =   255
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   645
         MaxLength       =   4
         TabIndex        =   473
         Text            =   "0"
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   645
         MaxLength       =   4
         TabIndex        =   472
         Text            =   "0"
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   645
         MaxLength       =   4
         TabIndex        =   471
         Text            =   "0"
         Top             =   1245
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   645
         MaxLength       =   4
         TabIndex        =   470
         Text            =   "0"
         Top             =   1575
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   645
         MaxLength       =   4
         TabIndex        =   469
         Text            =   "0"
         Top             =   1905
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   645
         MaxLength       =   4
         TabIndex        =   468
         Text            =   "0"
         Top             =   2235
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   645
         MaxLength       =   4
         TabIndex        =   467
         Text            =   "0"
         Top             =   2565
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   645
         MaxLength       =   4
         TabIndex        =   466
         Text            =   "0"
         Top             =   2895
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   645
         MaxLength       =   4
         TabIndex        =   465
         Text            =   "0"
         Top             =   3225
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   645
         MaxLength       =   4
         TabIndex        =   464
         Text            =   "0"
         Top             =   3555
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   645
         MaxLength       =   4
         TabIndex        =   463
         Text            =   "0"
         Top             =   3885
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   645
         MaxLength       =   4
         TabIndex        =   462
         Text            =   "0"
         Top             =   4215
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   645
         MaxLength       =   4
         TabIndex        =   461
         Text            =   "0"
         Top             =   4545
         Width           =   555
      End
      Begin VB.TextBox txtInsurance3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   645
         MaxLength       =   4
         TabIndex        =   460
         Text            =   "0"
         Top             =   4875
         Width           =   555
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   459
         Text            =   "0"
         Top             =   4875
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   458
         Text            =   "0"
         Top             =   4545
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   457
         Text            =   "0"
         Top             =   4215
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   456
         Text            =   "0"
         Top             =   3885
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   455
         Text            =   "0"
         Top             =   3555
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   454
         Text            =   "0"
         Top             =   3225
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   453
         Text            =   "0"
         Top             =   2895
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   452
         Text            =   "0"
         Top             =   2565
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   451
         Text            =   "0"
         Top             =   2235
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   450
         Text            =   "0"
         Top             =   1905
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   449
         Text            =   "0"
         Top             =   1575
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   448
         Text            =   "0"
         Top             =   1245
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   447
         Text            =   "0"
         Top             =   915
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   446
         Text            =   "0"
         Top             =   585
         Width           =   525
      End
      Begin VB.TextBox txtASF3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   445
         Text            =   "0"
         Top             =   255
         Width           =   525
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   444
         Text            =   "0"
         Top             =   4860
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   443
         Text            =   "0"
         Top             =   4530
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   442
         Text            =   "0"
         Top             =   4200
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   441
         Text            =   "0"
         Top             =   3870
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   440
         Text            =   "0"
         Top             =   3540
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   439
         Text            =   "0"
         Top             =   3210
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   438
         Text            =   "0"
         Top             =   2880
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   437
         Text            =   "0"
         Top             =   2550
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   436
         Text            =   "0"
         Top             =   2220
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   435
         Text            =   "0"
         Top             =   1890
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   434
         Text            =   "0"
         Top             =   1560
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   433
         Text            =   "0"
         Top             =   1230
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   432
         Text            =   "0"
         Top             =   900
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   431
         Text            =   "0"
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   430
         Text            =   "0"
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   429
         Text            =   "0"
         Top             =   4860
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   428
         Text            =   "0"
         Top             =   4530
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   427
         Text            =   "0"
         Top             =   4200
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   426
         Text            =   "0"
         Top             =   3870
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   425
         Text            =   "0"
         Top             =   3540
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   424
         Text            =   "0"
         Top             =   3210
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   423
         Text            =   "0"
         Top             =   2880
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   422
         Text            =   "0"
         Top             =   2550
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   421
         Text            =   "0"
         Top             =   2220
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   420
         Text            =   "0"
         Top             =   1890
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   419
         Text            =   "0"
         Top             =   1560
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   418
         Text            =   "0"
         Top             =   1230
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   417
         Text            =   "0"
         Top             =   900
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   416
         Text            =   "0"
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox txtMeals3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   415
         Text            =   "0"
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   5340
      Left            =   3630
      TabIndex        =   338
      Top             =   3855
      Visible         =   0   'False
      Width           =   3225
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   625
         Text            =   "0"
         Top             =   4860
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   624
         Text            =   "0"
         Top             =   4530
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   623
         Text            =   "0"
         Top             =   4200
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   622
         Text            =   "0"
         Top             =   3870
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   621
         Text            =   "0"
         Top             =   3540
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   620
         Text            =   "0"
         Top             =   3210
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   619
         Text            =   "0"
         Top             =   2880
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   618
         Text            =   "0"
         Top             =   2550
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   617
         Text            =   "0"
         Top             =   2220
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   616
         Text            =   "0"
         Top             =   1890
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   615
         Text            =   "0"
         Top             =   1560
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   614
         Text            =   "0"
         Top             =   1230
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   613
         Text            =   "0"
         Top             =   900
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   612
         Text            =   "0"
         Top             =   570
         Width           =   450
      End
      Begin VB.TextBox txtMisc_R2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   611
         Text            =   "0"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   105
         MaxLength       =   4
         TabIndex        =   413
         Text            =   "0"
         Top             =   4875
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   105
         MaxLength       =   4
         TabIndex        =   412
         Text            =   "0"
         Top             =   4545
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   105
         MaxLength       =   4
         TabIndex        =   411
         Text            =   "0"
         Top             =   4215
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   105
         MaxLength       =   4
         TabIndex        =   410
         Text            =   "0"
         Top             =   3885
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   105
         MaxLength       =   4
         TabIndex        =   409
         Text            =   "0"
         Top             =   3555
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   105
         MaxLength       =   4
         TabIndex        =   408
         Text            =   "0"
         Top             =   3225
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   105
         MaxLength       =   4
         TabIndex        =   407
         Text            =   "0"
         Top             =   2895
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   105
         MaxLength       =   4
         TabIndex        =   406
         Text            =   "0"
         Top             =   2565
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   105
         MaxLength       =   4
         TabIndex        =   405
         Text            =   "0"
         Top             =   2235
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   105
         MaxLength       =   4
         TabIndex        =   404
         Text            =   "0"
         Top             =   1905
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   105
         MaxLength       =   4
         TabIndex        =   403
         Text            =   "0"
         Top             =   1575
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   105
         MaxLength       =   4
         TabIndex        =   402
         Text            =   "0"
         Top             =   1245
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   105
         MaxLength       =   4
         TabIndex        =   401
         Text            =   "0"
         Top             =   915
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   105
         MaxLength       =   4
         TabIndex        =   400
         Text            =   "0"
         Top             =   585
         Width           =   510
      End
      Begin VB.TextBox txtCompaComm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   105
         MaxLength       =   4
         TabIndex        =   399
         Text            =   "0"
         Top             =   255
         Width           =   510
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   630
         MaxLength       =   4
         TabIndex        =   398
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   630
         MaxLength       =   4
         TabIndex        =   397
         Text            =   "0"
         Top             =   585
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   630
         MaxLength       =   4
         TabIndex        =   396
         Text            =   "0"
         Top             =   915
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   630
         MaxLength       =   4
         TabIndex        =   395
         Text            =   "0"
         Top             =   1245
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   630
         MaxLength       =   4
         TabIndex        =   394
         Text            =   "0"
         Top             =   1575
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   630
         MaxLength       =   4
         TabIndex        =   393
         Text            =   "0"
         Top             =   1905
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   630
         MaxLength       =   4
         TabIndex        =   392
         Text            =   "0"
         Top             =   2235
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   630
         MaxLength       =   4
         TabIndex        =   391
         Text            =   "0"
         Top             =   2565
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   630
         MaxLength       =   4
         TabIndex        =   390
         Text            =   "0"
         Top             =   2895
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   630
         MaxLength       =   4
         TabIndex        =   389
         Text            =   "0"
         Top             =   3225
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   630
         MaxLength       =   4
         TabIndex        =   388
         Text            =   "0"
         Top             =   3555
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   630
         MaxLength       =   4
         TabIndex        =   387
         Text            =   "0"
         Top             =   3885
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   630
         MaxLength       =   4
         TabIndex        =   386
         Text            =   "0"
         Top             =   4215
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   630
         MaxLength       =   4
         TabIndex        =   385
         Text            =   "0"
         Top             =   4545
         Width           =   480
      End
      Begin VB.TextBox txtInsurance2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   630
         MaxLength       =   4
         TabIndex        =   384
         Text            =   "0"
         Top             =   4875
         Width           =   480
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   383
         Text            =   "0"
         Top             =   4875
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   382
         Text            =   "0"
         Top             =   4545
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   381
         Text            =   "0"
         Top             =   4215
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   380
         Text            =   "0"
         Top             =   3885
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   379
         Text            =   "0"
         Top             =   3555
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   378
         Text            =   "0"
         Top             =   3225
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   377
         Text            =   "0"
         Top             =   2895
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   376
         Text            =   "0"
         Top             =   2565
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   375
         Text            =   "0"
         Top             =   2235
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   374
         Text            =   "0"
         Top             =   1905
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   373
         Text            =   "0"
         Top             =   1575
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   372
         Text            =   "0"
         Top             =   1245
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   371
         Text            =   "0"
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   370
         Text            =   "0"
         Top             =   585
         Width           =   495
      End
      Begin VB.TextBox txtASF2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   369
         Text            =   "0"
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   368
         Text            =   "0"
         Top             =   4860
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   367
         Text            =   "0"
         Top             =   4530
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   366
         Text            =   "0"
         Top             =   4200
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   365
         Text            =   "0"
         Top             =   3870
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   364
         Text            =   "0"
         Top             =   3540
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   363
         Text            =   "0"
         Top             =   3210
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   362
         Text            =   "0"
         Top             =   2880
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   361
         Text            =   "0"
         Top             =   2550
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   360
         Text            =   "0"
         Top             =   2220
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   359
         Text            =   "0"
         Top             =   1890
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   358
         Text            =   "0"
         Top             =   1560
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   357
         Text            =   "0"
         Top             =   1230
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   356
         Text            =   "0"
         Top             =   900
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   355
         Text            =   "0"
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox txtTerminalFee2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   354
         Text            =   "0"
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   353
         Text            =   "0"
         Top             =   4860
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   13
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   352
         Text            =   "0"
         Top             =   4530
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   351
         Text            =   "0"
         Top             =   4200
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   350
         Text            =   "0"
         Top             =   3870
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   349
         Text            =   "0"
         Top             =   3540
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   348
         Text            =   "0"
         Top             =   3210
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   347
         Text            =   "0"
         Top             =   2880
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   346
         Text            =   "0"
         Top             =   2550
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   345
         Text            =   "0"
         Top             =   2220
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   344
         Text            =   "0"
         Top             =   1890
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   343
         Text            =   "0"
         Top             =   1560
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   342
         Text            =   "0"
         Top             =   1230
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   341
         Text            =   "0"
         Top             =   900
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   340
         Text            =   "0"
         Top             =   570
         Width           =   465
      End
      Begin VB.TextBox txtMeals2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   339
         Text            =   "0"
         Top             =   240
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SQL As String
Dim Commission As Double
Dim Discount As Double
Dim RsComm As ADODB.Recordset
Dim RsVAT As ADODB.Recordset

Dim CompanyComm(1 To 4) As Double
Dim Insurance(1 To 4) As Double
Dim ASF(1 To 4) As Double
Dim TerminalFee(1 To 4) As Double
Dim Meals(1 To 4) As Double
Dim Misc(1 To 4) As Double


Dim myArr_Evat(0 To 14) As Double
Dim myArr_ServiceFee(0 To 14) As Double
Dim myArr_RefundFee(0 To 14) As Double
Dim myArr_VoidFee(0 To 14) As Double
Dim myArr_NoShowFee(0 To 14) As Double



Dim myRouteEVAT1(0 To 14) As Double
Dim myRouteEVAT2(0 To 14) As Double
Dim myRouteEVAT3(0 To 14) As Double
Dim myRouteEVAT4(0 To 14) As Double

Dim myRouteServiceFEE1(0 To 14) As Double
Dim myRouteServiceFEE2(0 To 14) As Double
Dim myRouteServiceFEE3(0 To 14) As Double
Dim myRouteServiceFEE4(0 To 14) As Double

Dim myRouteRefundFEE1(0 To 14) As Double
Dim myRouteRefundFEE2(0 To 14) As Double
Dim myRouteRefundFEE3(0 To 14) As Double
Dim myRouteRefundFEE4(0 To 14) As Double

Dim myRouteVoidFEE1(0 To 14) As Double
Dim myRouteVoidFEE2(0 To 14) As Double
Dim myRouteVoidFEE3(0 To 14) As Double
Dim myRouteVoidFEE4(0 To 14) As Double

Dim myRouteNoShowFEE1(0 To 14) As Double
Dim myRouteNoShowFEE2(0 To 14) As Double
Dim myRouteNoShowFEE3(0 To 14) As Double
Dim myRouteNoShowFEE4(0 To 14) As Double



Private Sub CboAccountName_Click()
If Not CheckNull(Me.CboAccountName) Then
     txtAgencyName = FindAccountName(Me.CboAccountName)
     Me.txtCommPercent = Format(ReturnAccDetails(FindAirline(Me.Combo1), FindAccountID(CboAccountName), "Commission"), "###,##0.00")
End If
End Sub

Function ReturnAcc_Amount(ByVal UserFld As String, ByVal UserCri As String) As Double
On Error GoTo FailSafe_Error
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_CustAccounts WHERE [Account Name]='" & UCase(UserCri) & "'"
With Rst

        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              ReturnAcc_Amount = .Fields(UserFld).Value
          Else
              ReturnAcc_Amount = 0
        End If
     .Close
   Set Rst = Nothing
End With
Exit Function
FailSafe_Error:
ReturnAcc_Amount = 0
End Function

Function FindAccountName(Param) As String
Dim Rst As New ADODB.Recordset

SQL = "SELECT * FROM tbl_CustAccounts WHERE [Account Name]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindAccountName = .Fields(1).Value
          Else
              FindAccountName = ""
        End If
     .Close
   Set Rst = Nothing
End With
End Function

Function FindAccountID(Param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_CustAccounts WHERE [Account Name]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindAccountID = .Fields(0).Value
          Else
              FindAccountID = -1
        End If
     .Close
   Set Rst = Nothing
End With
End Function


Private Sub cboDest1_Change(Index As Integer)
            Call Recalc(Index)
            If cboDest1(Index).Text = "" Then
           
                                                    txtCompaComm2(Index).Text = "0.00"
                                                    Me.txtInsurance2(Index).Text = "0.00"
                                                    Me.txtASF2(Index).Text = "0.00"
                                                    Me.txtTerminalFee2(Index).Text = "0.00"
                                                    Me.txtMeals2(Index).Text = "0.00"
                                                    Me.txtMisc_R2(Index).Text = "0.00"
           
           End If
End Sub

Private Sub cboDest1_Click(Index As Integer)
   Call Recalc(Index)
If Me.Check1.Value = 1 Then
    Dim ask As Integer
    ask = MsgBox("Add Meal to this ticket?", vbInformation + vbYesNo)
    If ask = vbYes Then
        frmMeals.Tag = Index
        frmMeals.Show 1
        Exit Sub
    End If
End If
End Sub



Private Sub cboDest1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cboDest2_Change(Index As Integer)
           Call Recalc(Index)
           If cboDest2(Index).Text = "" Then
           
                                                    txtCompaComm3(Index).Text = "0.00"
                                                    Me.txtInsurance3(Index).Text = "0.00"
                                                    Me.txtASF3(Index).Text = "0.00"
                                                    Me.txtTerminalFee3(Index).Text = "0.00"
                                                    Me.txtMeals3(Index).Text = "0.00"
                                                    Me.txtMisc_R3(Index).Text = "0.00"
           
           End If
End Sub

Private Sub cboDest2_Click(Index As Integer)
'Dim i As Integer
'If Index = 0 Then
'===========================
        'For i = 1 To 14
        '    Me.cboDest2(i).Text = Me.cboDest2(Index).Text
        '   Call ReCalc(0)
        '   Call ReCalc(i)
        '   DoEvents
        'Next i
'        Me.cboDest2(1).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(2).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(3).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(4).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(5).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(6).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(7).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(8).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(9).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(10).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(11).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(12).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(13).Text = Me.cboDest2(Index).Text
'        Me.cboDest2(14).Text = Me.cboDest2(Index).Text
'        Call ReCalc(0)
'        Call ReCalc(1)
'        Call ReCalc(2)
'        Call ReCalc(3)
'        Call ReCalc(4)
'        Call ReCalc(5)
'        Call ReCalc(6)
'        Call ReCalc(7)
'        Call ReCalc(8)
'        Call ReCalc(9)
'        Call ReCalc(10)
'        Call ReCalc(11)
'        Call ReCalc(12)
'        Call ReCalc(13)
'        Call ReCalc(14)

        
'Else
            Call Recalc(Index)
        
'End If



End Sub

Private Sub cboDest2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cboDest3_Change(Index As Integer)
           Call Recalc(Index)
           If cboDest3(Index).Text = "" Then
           
                                                    txtCompaComm4(Index).Text = "0.00"
                                                    Me.txtInsurance4(Index).Text = "0.00"
                                                    Me.txtASF4(Index).Text = "0.00"
                                                    Me.txtTerminalFee4(Index).Text = "0.00"
                                                    Me.txtMeals4(Index).Text = "0.00"
                                                    Me.txtMisc_R4(Index).Text = "0.00"
           
           End If
End Sub

Private Sub cboDest3_Click(Index As Integer)
Call Recalc(Index)
End Sub

Private Sub cboDest3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cboFrom_Change(Index As Integer)

            Call Recalc(Index)
           If cboFrom(Index).Text = "" Then
           
                                                    txtCompaComm1(Index).Text = "0.00"
                                                    Me.txtInsurance1(Index).Text = "0.00"
                                                    Me.txtASF1(Index).Text = "0.00"
                                                    Me.txtTerminalFee1(Index).Text = "0.00"
                                                    Me.txtMeals1(Index).Text = "0.00"
                                                    Me.txtMisc_R1(Index).Text = "0.00"
           
           End If
End Sub

Private Sub cboFrom_Click(Index As Integer)
   Call Recalc(Index)
If Me.Check1.Value = 1 Then
    Dim ask As Integer
    ask = MsgBox("Add Meal to this ticket?", vbInformation + vbYesNo)
    If ask = vbYes Then
        frmMeals.Tag = Index
        frmMeals.Show 1
        Exit Sub
    End If
End If

End Sub

Private Sub cboFrom_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cboPassengerType_Click(Index As Integer)
            txtDiscountAmt(Index).Text = FindPassengerType(Me.cboPassengerType(Index))
            Call Recalc(Index)

End Sub

Private Sub cboPassengerType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub




Private Sub cboTicketType_GotFocus(Index As Integer)
cboTicketType(Index).Width = 1500
End Sub

Private Sub cboTicketType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cboTicketType_LostFocus(Index As Integer)
cboTicketType(Index).Width = 750
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Call Clear
Me.Caption = "new"

If CheckNull(Me.Combo1) Then
    MsgBox "Please select an airline/shipping line"
Else
Me.Frame2.Enabled = True
frmSelect.Show 1
End If

End Sub

Sub Clear()
Dim i As Integer

    Me.txtCommPercent = 0
    Me.txtNo = ""
For i = 0 To 14 Step 1
    txtTicketNo(i).Text = Empty
    Me.txtPassengerName(i).Text = Empty
    Me.cboFrom(i).Text = Empty
    Me.cboDest1(i).Text = Empty
    Me.cboDest2(i).Text = Empty
    Me.cboDest3(i).Text = Empty
    
    Me.txtDepDate1(i).Text = Empty
    Me.txtDepDate2(i).Text = Empty
    Me.txtDepDate3(i).Text = Empty
    
    Me.cboTicketType(i).Text = Empty
    Me.cboPassengerType(i).Text = Empty
    Me.txtNet(i).Text = "0.00"
    Me.txtDiscountAmt(i).Text = "0.00"
    Me.txtGrossFare(i).Text = "0.00"
    Me.txtFare(i).Text = "0.00"
    
    Me.txtCompaComm1(i).Text = "0.00"
    Me.txtInsurance1(i).Text = "0.00"
    Me.txtASF1(i).Text = "0.00"
    Me.txtTerminalFee1(i).Text = "0.00"
    Me.txtMeals1(i).Text = "0.00"
    Me.txtMisc_R1(i).Text = "0.00"
    
    
    Me.txtCompaComm2(i).Text = "0.00"
    Me.txtInsurance2(i).Text = "0.00"
    Me.txtASF2(i).Text = "0.00"
    Me.txtTerminalFee2(i).Text = "0.00"
    Me.txtMeals2(i).Text = "0.00"
    Me.txtMisc_R2(i).Text = "0.00"

    Me.txtCompaComm3(i).Text = "0.00"
    Me.txtInsurance3(i).Text = "0.00"
    Me.txtASF3(i).Text = "0.00"
    Me.txtTerminalFee3(i).Text = "0.00"
    Me.txtMeals3(i).Text = "0.00"
    Me.txtMisc_R3(i).Text = "0.00"


    Me.txtCompaComm4(i).Text = "0.00"
    Me.txtInsurance4(i).Text = "0.00"
    Me.txtASF4(i).Text = "0.00"
    Me.txtTerminalFee4(i).Text = "0.00"
    Me.txtMeals4(i).Text = "0.00"
    Me.txtMisc_R4(i).Text = "0.00"




    Me.txtMisc = "0.00"
    myArr_Evat(i) = 0
    myArr_ServiceFee(i) = 0
    myArr_RefundFee(i) = 0
    
    
     myRouteEVAT1(i) = 0
     myRouteEVAT2(i) = 0
     myRouteEVAT3(i) = 0
     myRouteEVAT4(i) = 0
    
     myRouteServiceFEE1(i) = 0
     myRouteServiceFEE2(i) = 0
     myRouteServiceFEE3(i) = 0
     myRouteServiceFEE4(i) = 0
    
     myRouteRefundFEE1(i) = 0
     myRouteRefundFEE2(i) = 0
     myRouteRefundFEE3(i) = 0
     myRouteRefundFEE4(i) = 0
     
     myRouteVoidFEE1(i) = 0
     myRouteVoidFEE2(i) = 0
     myRouteVoidFEE3(i) = 0
     myRouteVoidFEE4(i) = 0
        
     myRouteNoShowFEE1(i) = 0
     myRouteNoShowFEE2(i) = 0
     myRouteNoShowFEE3(i) = 0
     myRouteNoShowFEE4(i) = 0
        
     myArr_VoidFee(0) = 0
     myArr_NoShowFee(0) = 0
        
     Me.txtTotFare = "0.00"
     txtTotDiscount = "0.00"
     txtTotComm = "0.00"
     txtMisc = "0.00"
     Me.CboAccountName = ""
     txtComm(i).Text = CDbl(txtCommPercent)
     Recalc (i)
Next i
Me.txtCommPercent.Enabled = False
End Sub

Private Sub cmdOverRide_Click()
Me.Frame2.Enabled = True
Me.Tag = "over_ride"
frmSelectStatement.Tag = "over_ride"
frmSelectStatement.Show 1
End Sub

Function Remove_SA(ByVal Param As String) As Boolean
On Error GoTo FailSafe_Error
Dim RsDel As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & Param & "'"
With RsDel
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        SQL = "DELETE * FROM tbl_Statement WHERE [sNumber]='" & Param & "'"
            cn.BeginTrans
                    cn.Execute SQL
            cn.CommitTrans
            Remove_SA = True
        Else
            Remove_SA = False
        End If
End With
Exit Function
FailSafe_Error:
cn.RollbackTrans
End Function

Private Sub cmdPost_Click()
'On Error GoTo ErrExit
Dim RsPost As ADODB.Recordset
Dim RsPostDetail As ADODB.Recordset
Dim RsStatementTickets As ADODB.Recordset
Dim ask As Integer
Dim i, j As Integer
Dim myTempCp(1 To 4) As Variant

If CheckNull(Me.CboAccountName) Then: MsgBox "Account name should not be blank": Exit Sub
If CheckNull(txtAgencyName) Then: MsgBox "Agency name should not be blank": Exit Sub



Set RsPost = New ADODB.Recordset
Set RsPostDetail = New ADODB.Recordset
Set RsStatementTickets = New ADODB.Recordset


RsPostDetail.Open "SELECT * FROM tbl_StatementDetail", cn, adOpenKeyset, adLockOptimistic
RsStatementTickets.Open "SELECT * FROM tbl_StatementTickets", cn, adOpenKeyset, adLockOptimistic

SQL = "SELECT * FROM tbl_Statement"
ask = MsgBox("Are all the information correct?", vbYesNo + vbInformation, "ELS")

If ask = vbYes Then
 

If Remove_SA(Me.txtNo) Then
            MsgBox "Previous statement was updated...", vbInformation
Else
    If ReturnFirst(frmStatement.GetLastNumber) = 0 Then
            'frmStatement.txtNo = AutoIncrement(frmStatement.GetLastNumber) & "-" & kulotRead(App.Path & "\Settings.txt")
            frmStatement.txtNo = frmStatement.GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
      Else
            'frmStatement.txtNo = AutoIncrement(Mid(frmStatement.GetLastNumber, 1, ReturnFirst(frmStatement.GetLastNumber))) & kulotRead(App.Path & "\Settings.txt")
            frmStatement.txtNo = frmStatement.GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
    End If
End If

    With RsPost
                .Open SQL, cn, adOpenKeyset, adLockOptimistic
                cn.BeginTrans
                .AddNew
                .Fields("sNumber").Value = Me.txtNo
                .Fields("AccountNo").Value = Me.CboAccountName
                .Fields("AgencyName").Value = Me.txtAgencyName
                .Fields("Date").Value = Me.txtDate
                .Fields("Total Amount").Value = CDbl(Me.txtTotFare)
                .Fields("Airline").Value = FindAirline(Me.Combo1)
                .Fields("Branch Number").Value = WhichBranch.Fields(2).Value
                .Fields("Paid").Value = False
                 
                .Fields("Void").Value = False
                .Fields("Refund").Value = False
                .Fields("Balance").Value = CDbl(Me.txtTotFare)
                .Fields("Credit Card Activated").Value = False
                .Fields("AccountID").Value = FindAccountID(Me.CboAccountName)
                .Fields("MiscAmt").Value = CDbl(Me.txtMisc)
                .Fields("Issued by").Value = Me.txtIssuedBy
                .Update
                Me.Caption = "new"
                cn.CommitTrans
                
                        With RsPostDetail
                                    For i = 0 To 14 Step 1
                                      If IsNumeric(Me.txtTicketNo(i)) Then
                                       cn.BeginTrans
                                            .AddNew
                                            .Fields("TransID").Value = RsPost.Fields(0).Value
                                            .Fields("Ticket No").Value = Me.txtTicketNo(i)
                                            .Fields("Name").Value = Me.txtPassengerName(i)
                                            .Fields("Ticket Type").Value = Me.cboTicketType(i)
                                            .Fields("Passenger Type").Value = Me.cboPassengerType(i)
                                            .Fields("Gross").Value = Me.txtGrossFare(i)
                                            .Fields("Commision").Value = Me.txtComm(i)
                                            .Fields("Net").Value = Me.txtNet(i)
                                            .Fields("Discount").Value = Me.txtDiscountAmt(i)
                                            .Fields("Fare").Value = Me.txtFare(i)
                                            .Fields("Airline").Value = FindAirline(Me.Combo1)
                                            .Fields("Paid").Value = False
                                            
                                            If Me.cboPassengerType(i).Text <> "VOID" Then
                                                .Fields("Void").Value = False
                                            Else
                                                .Fields("Void").Value = True
                                            End If
                                            
                                            .Fields("Refund").Value = False
                                            .Fields("EVAT").Value = myArr_Evat(i)
                                            .Fields("Service Fee").Value = myArr_ServiceFee(i)
                                            .Fields("Refund Fee").Value = myArr_RefundFee(i)
                                            .Fields("Void Fee").Value = myArr_VoidFee(i)
                                            .Fields("Noshow Fee").Value = myArr_NoShowFee(i)
                                            Call Check_Update_Ticket(Me.txtTicketNo(i), "Edit")
                                            .Update
                                            
'=========

myTempCp(1) = Mid$(Me.cboFrom(i), ReturnPos(Me.cboFrom(i)) + 1, Len(Me.cboFrom(i)))
myTempCp(2) = Mid$(Me.cboDest1(i), ReturnPos(Me.cboDest1(i)) + 1, Len(Me.cboDest1(i)))
myTempCp(3) = Mid$(Me.cboDest2(i), ReturnPos(Me.cboDest2(i)) + 1, Len(Me.cboDest2(i)))
myTempCp(4) = Mid$(Me.cboDest3(i), ReturnPos(Me.cboDest3(i)) + 1, Len(Me.cboDest3(i)))

                                       For j = 1 To 4 Step 1
                                           If j = 1 Then
                                           If Me.cboFrom(i).Text <> "" Then
                                           With RsStatementTickets
                                           .AddNew
                                           .Fields("StatementDetails").Value = RsPostDetail.Fields(0).Value
                                           .Fields("DepartDateTime").Value = Me.txtDepDate1(i)
                                           .Fields("Route").Value = Me.cboFrom(i)
                                           
                                           .Fields("Company Commission").Value = Me.txtCompaComm1(i)
                                           .Fields("Insurance").Value = Me.txtInsurance1(i)
                                           .Fields("ASF").Value = Me.txtASF1(i)
                                           .Fields("Terminal Fee").Value = Me.txtTerminalFee1(i)
                                           .Fields("Meals").Value = Me.txtMeals1(i)
                                           .Fields("Misc").Value = Me.txtMisc_R1(i)
                                           .Fields("Refund").Value = False
                                           


                                           '.Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboFrom(i), i, Mid$(Me.cboFrom(Index), 8, 3)))
                                           .Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboFrom(i), i, myTempCp(1)))
                                           .Fields("VAT").Value = ReturnVAT()
                                           .Update
                                           End With
                                           End If
                                           End If
                                           
                                           If j = 2 Then
                                           If Me.cboDest1(i).Text <> "" Then
                                           With RsStatementTickets
                                           .AddNew
                                           .Fields("StatementDetails").Value = RsPostDetail.Fields(0).Value
                                           .Fields("DepartDateTime").Value = Me.txtDepDate2(i)
                                           .Fields("Route").Value = Me.cboDest1(i)
                                           
                                           .Fields("Company Commission").Value = Me.txtCompaComm2(i)
                                           .Fields("Insurance").Value = Me.txtInsurance2(i)
                                           .Fields("ASF").Value = Me.txtASF2(i)
                                           .Fields("Terminal Fee").Value = Me.txtTerminalFee2(i)
                                           .Fields("Meals").Value = Me.txtMeals2(i)
                                           .Fields("Refund").Value = False
                                           .Fields("Misc").Value = Me.txtMisc_R2(i)
                                           '.Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboDest1(i), i, Mid$(Me.cboDest1(Index), 8, 3)))
                                           .Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboDest1(i), i, myTempCp(2)))
                                           .Fields("VAT").Value = ReturnVAT()
                                           .Update
                                           End With
                                           End If
                                           End If
                                           
                                           If j = 3 Then
                                           If Me.cboDest2(i).Text <> "" Then
                                           With RsStatementTickets
                                           .AddNew
                                           .Fields("StatementDetails").Value = RsPostDetail.Fields(0).Value
                                           .Fields("DepartDateTime").Value = Me.txtDepDate3(i)
                                           .Fields("Route").Value = Me.cboDest2(i)
                                           
                                           .Fields("Company Commission").Value = Me.txtCompaComm3(i)
                                           .Fields("Insurance").Value = Me.txtInsurance3(i)
                                           .Fields("ASF").Value = Me.txtASF3(i)
                                           .Fields("Terminal Fee").Value = Me.txtTerminalFee3(i)
                                           .Fields("Meals").Value = Me.txtMeals3(i)
                                           .Fields("Refund").Value = False
                                           .Fields("Misc").Value = Me.txtMisc_R3(i)
                                           '.Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboDest2(i), i, Mid$(Me.cboDest2(Index), 8, 3)))
                                           .Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboDest2(i), i, myTempCp(3)))
                                           .Fields("VAT").Value = ReturnVAT()
                                           .Update
                                           End With
                                           End If
                                           End If
                                           
                                           
                                           If j = 4 Then
                                           If Me.cboDest3(i).Text <> "" Then
                                           With RsStatementTickets
                                           .AddNew
                                           .Fields("StatementDetails").Value = RsPostDetail.Fields(0).Value
                                           .Fields("DepartDateTime").Value = Me.txtDepDate3(i) ' 3 LANG ANAY KAY WLA LUGAR
                                           .Fields("Route").Value = Me.cboDest3(i)
                                           
                                           .Fields("Company Commission").Value = Me.txtCompaComm4(i)
                                           .Fields("Insurance").Value = Me.txtInsurance4(i)
                                           .Fields("ASF").Value = Me.txtASF4(i)
                                           .Fields("Terminal Fee").Value = Me.txtTerminalFee4(i)
                                           .Fields("Meals").Value = Me.txtMeals4(i)
                                           .Fields("Refund").Value = False
                                           .Fields("Misc").Value = Me.txtMisc_R4(i)
                                           '.Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboDest3(i), i, Mid$(Me.cboDest3(Index), 8, 3)))
                                           .Fields("TicketAmount").Value = CDbl(GetRouteAmt(Me.cboDest3(i), i, myTempCp(4)))
                                           .Fields("VAT").Value = ReturnVAT()
                                           .Update
                                           End With
                                           End If
                                           End If

                                           
                                       Next j
                                  
                                        cn.CommitTrans
                                      End If
                                  Next i
                        End With
            

    End With
    
                    
                    Dim AskPrint As Integer
                    Dim Rpt As New RptStatement
                    AskPrint = MsgBox("PLEASE INSERT PAPER AND CLICK OK TO START PRINTING...", vbOKCancel + vbExclamation)
                    If AskPrint = vbOK Then
                    
                               With Rpt
                                    .DataControl1.Connection = cn
                                    .DataControl1.Source = "SELECT * FROM qryStatement WHERE [sNumber]='" & Me.txtNo & "' ORDER by [Ticket No]"
                                    .Show 1
                                End With
                                Set Rpt = Nothing
                                Call Mark_As_Printed(Me.txtNo)
                    End If
                    Call Clear
Else
                    MsgBox "Operation cancelled..."
End If
Me.cmdPost.Enabled = False
Me.cmdRecalc.Enabled = True

                    
Exit Sub
ErrExit:
cn.RollbackTrans
MsgBox "There was an error while trying to save the statement pls try again."
End Sub

Function Mark_As_Printed(Param)
On Error GoTo FailSafe_Err
Dim Rst     As New ADODB.Recordset
Dim mySQL   As String

mySQL = "UPDATE tbl_Statement SET tbl_Statement.Printed = True " & _
        "WHERE (((tbl_Statement.sNumber)='" & Param & "'));"

With cn
            .BeginTrans
                .Execute mySQL
            .CommitTrans
End With
    Exit Function
FailSafe_Err:
cn.RollbackTrans
End Function


Private Sub cmdRecalc_Click()

Dim myTmpBalancePeso        As Double
Dim myTmpBalanceDollar      As Double
Dim myTmpLimitPeso          As Double
Dim mytmpLimitDollar        As Double
Dim myTmpFare               As Double
Dim myTotalBalance          As Double
Dim i As Integer

For i = 0 To 14
    txtComm(i).Text = CDbl(txtCommPercent)
    Recalc (i)
Next i
'Call cmdSet_Click


If Me.Caption = "new" Then
        If ReturnFirst(frmStatement.GetLastNumber) = 0 Then
            frmStatement.txtNo = frmStatement.GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
        Else
            'frmStatement.txtNo = AutoIncrement(Mid(frmStatement.GetLastNumber, 1, ReturnFirst(frmStatement.GetLastNumber))) & kulotRead(App.Path & "\Settings.txt")
            frmStatement.txtNo = frmStatement.GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
        End If
End If

   myTmpBalancePeso = ReturnAcc_Amount("Current Balance Peso", Me.CboAccountName)
   myTmpBalanceDollar = ReturnAcc_Amount("Current Balance Dollar", Me.CboAccountName)
   myTmpLimitPeso = ReturnAcc_Amount("Credit Limit Peso", Me.CboAccountName)
   mytmpLimitDollar = ReturnAcc_Amount("Credit Limit Dollar", Me.CboAccountName)
   
   myTmpFare = FareCounter + CDbl(Me.txtMisc)
   
   myTotalBalance = myTmpFare + myTmpBalancePeso
   
   
   If myTotalBalance > myTmpLimitPeso Then
   Dim ask As Integer
   MsgBox "Warning!!! This account name :" & Me.CboAccountName & " exceeds Peso credit limit! this statement will not be posted!", vbInformation
   ask = MsgBox("Over-ride", vbYesNo + vbCritical)
   
   If ask = vbYes Then
    GoTo Cont
   End If
    Exit Sub
   End If
Cont:
Me.txtTotFare = Format(myTmpFare, "###,##0.00")
Me.txtTotComm = Format(Commission, "###,##0.00")
Me.txtTotDiscount = Format(Discount, "###,##0.00")
Me.cmdPost.Enabled = True
Me.Tag = ""
'Me.cmdRecalc.Enabled = False
End Sub

Function FareCounter() As Double
Dim i As Integer
Dim TmpCtr As Double
TmpCtr = 0
Commission = 0
Discount = 0
For i = 0 To 14 Step 1
    If IsNumeric(Me.txtTicketNo(i)) Then
            If IsNumeric(Me.txtFare(i)) Then
                TmpCtr = CDbl(TmpCtr) + CDbl(Me.txtFare(i))
                
                Commission = CDbl(Commission) + (CDbl(txtGrossFare(i)) * (CDbl(txtComm(i)) / 100))
                Discount = CDbl(Discount) + (CDbl(txtGrossFare(i)) * (CDbl(txtDiscountAmt(i) / 100)))
            End If
    End If
Next i
Call EraseExcess
'MsgBox TmpCtr
FareCounter = TmpCtr
End Function

Sub EraseExcess()
Dim i As Integer
    For i = 0 To 14
        If Not IsNumeric(Me.txtTicketNo(i)) Then
            Me.cboFrom(i).Text = ""
            Me.cboDest1(i).Text = ""
            Me.cboDest2(i).Text = ""
            txtDepDate1(i).Text = ""
            txtDepDate2(i).Text = ""
            txtDepDate3(i).Text = ""
            cboTicketType(i).Text = ""
            cboPassengerType(i).Text = ""
            txtGrossFare(i).Text = ""
            txtComm(i).Text = ""
            txtDiscountAmt(i).Text = ""
            txtFare(i).Text = ""
            Me.txtNet(i).Text = ""
            
            
    Me.txtCompaComm1(i).Text = ""
    Me.txtInsurance1(i).Text = ""
    Me.txtASF1(i).Text = ""
    Me.txtTerminalFee1(i).Text = ""
    Me.txtMeals1(i).Text = ""
    
    Me.txtCompaComm2(i).Text = ""
    Me.txtInsurance2(i).Text = ""
    Me.txtASF2(i).Text = ""
    Me.txtTerminalFee2(i).Text = ""
    Me.txtMeals2(i).Text = ""


    Me.txtCompaComm3(i).Text = ""
    Me.txtInsurance3(i).Text = ""
    Me.txtASF3(i).Text = ""
    Me.txtTerminalFee3(i).Text = ""
    Me.txtMeals3(i).Text = ""

            
            
            
            'Me.txtCompaComm(i).Text = ""
            'Me.txtInsurance(i).Text = ""
        End If
    Next i
End Sub

Private Sub cmdSet_Click()
frmUserVerify.Tag = "comm"
frmUserVerify.Show 1
'On Error GoTo FailSafe_Error
'cn.BeginTrans
'    RsComm.Update "Comm", CDbl(txtCommPercent)
'    RsComm.Requery
'cn.CommitTrans

'cn.RollbackTrans
'MsgBox "Commision cannot be save at this time please retry", vbCritical

End Sub

Private Sub Combo1_Click()
'Call FillRoutes
Call FillComboTicket
Me.txtCommPercent = Format(ReturnAccDetails(FindAirline(Me.Combo1), FindAccountID(CboAccountName), "Commission"), "###,##0.00")
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub





Private Sub Form_Activate()
MDImain.Arrange 2
End Sub

Function ReturnVAT() As Double
Set RsVAT = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Computation"
With RsVAT
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                ReturnVAT = .Fields(2).Value
            Else
                ReturnVAT = 0
        End If
        .Close
      Set RsVAT = Nothing
End With
End Function



Private Sub Form_Load()

'Set RsComm = New ADODB.Recordset
'SQL = "SELECT * FROM tbl_SetComm"
'With RsComm
'        .Open SQL, cn, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            txtCommPercent = .Fields(0).Value
'        Else
'            .AddNew
'            .Fields(0).Value = 0
'            .Update
'            .Requery
'            txtCommPercent = 0
'        End If
'End With

Me.txtDate = Format(Now, "mm/dd/yyyy")

Dim i As Integer

For i = 0 To 14
    txtComm(i).Text = CDbl(txtCommPercent)
    Recalc (i)
Next i


'Call cmdSet_Click
Call FillCombo
'Call FillRoutes
Call FillComboTicket
Call FillAccountName
Call Clear

If UCase(MDImain.StatusBar1.Panels(3).Text) <> "ADMIN" Then
    For i = 0 To 14
        Me.txtFare(i).Locked = True
        Me.txtGrossFare(i).Locked = True
    Next i
End If
Me.txtIssuedBy = MDImain.StatusBar1.Panels(2).Text
End Sub


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
Me.Combo1.ListIndex = 0

For i = 0 To 14
    Me.cboPassengerType(i).Clear
    Me.cboPassengerType(i).Enabled = False
Next i


End Sub

Sub FillAccountName()
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_CustAccounts ORDER BY [Account Name] ASC"
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    Do While Not .EOF
                    Me.CboAccountName.AddItem .Fields(1).Value
                    .MoveNext
                    Loop
               .Close
         Set Rst = Nothing
                 '  -> Me.CboAccountName.ListIndex = 0
            End If
End With

End Sub

Function FindPassengerType(Param) As String
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_PassengerType WHERE [PassengerType]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindPassengerType = .Fields(2).Value
          Else
              FindPassengerType = "0.00"
        End If
     .Close
   Set Rst = Nothing
End With
End Function

Sub FillRoutes(ByVal nTicketTypeID As Long)
Dim Tmp As ADODB.Recordset
Dim i As Integer
Set Tmp = New ADODB.Recordset
'01600244601
'SQL = "SELECT DISTINCT [FROM],[TO] FROM qryRoutePricing WHERE [AirlineID]=" & FindAirline(Me.Combo1) & " AND [TicketTypeID]=" & nTicketTypeID & " ORDER by [From] ASC"
SQL = "SELECT DISTINCT [FROM],[TO] FROM qryRoutePricing WHERE [AirlineID]=" & FindAirline(Me.Combo1) & " ORDER by [From] ASC"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
        For i = 0 To 14
            Me.cboFrom(i).AddItem .Fields(0).Value & "/" & .Fields(1).Value
            Me.cboDest1(i).AddItem .Fields(0).Value & "/" & .Fields(1).Value
            Me.cboDest2(i).AddItem .Fields(0).Value & "/" & .Fields(1).Value
            Me.cboDest3(i).AddItem .Fields(0).Value & "/" & .Fields(1).Value
        Next i
        .MoveNext
        Loop
    End If
End With
End Sub

Function ReturnIfAirline(Param) As String
Dim Tmp As New ADODB.Recordset

SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Me.Combo1 & "'"

With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    
    If .RecordCount > 0 Then
            ReturnIfAirline = .Fields(2).Value
          Else
            ReturnIfAirline = "n/a"
    End If
    .Close
End With
Set Tmp = Nothing
End Function

Sub FillComboTicket()
Dim Tmp As ADODB.Recordset
Dim i As Integer
Set Tmp = New ADODB.Recordset

SQL = "SELECT * FROM tbl_TicketType WHERE [AirlineShippingLine]='" & ReturnIfAirline(Me.Combo1) & "'"
'SQL = "SELECT * FROM qryStatementTicketType WHERE [AirlineID]=" & FindAirline(Me.Combo1)

For i = 0 To 14
    Me.cboTicketType(i).Clear
Next i

With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
    
        Do While Not .EOF
        For i = 0 To 14
            'Me.cboTicketType(i).AddItem .Fields(1).Value
            Me.cboTicketType(i).AddItem .Fields(1).Value
        Next i
        .MoveNext
        Loop
    End If
End With
End Sub

Function FindAirline(Param) As Long
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

Function FindTicket(Index) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_TicketType WHERE [Ticket Type]='" & cboTicketType(Index) & "'"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindTicket = .Fields(0).Value
      Else
        FindTicket = -1
    End If
    .Close
End With
Set Tmp = Nothing
End Function

Function GetBranchID() As Long
Dim Rst As New ADODB.Recordset
Dim SQL As String
SQL = "SELECT * FROM tbl_SetBranch"
Rst.Open SQL, cn, adOpenKeyset, adLockOptimistic
With Rst
If .RecordCount > 0 Then
    GetBranchID = .Fields(2).Value
End If
    .Close
End With
Set Rst = Nothing
End Function

Function GetLastNumber() As String
Dim RsFnumber       As ADODB.Recordset
Dim SQL             As String
Dim Tmp             As String
Dim myTmpPos        As Integer

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT sNumber from tbl_Statement ORDER by [sNumber] ASC"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               
               'GetLastNumber = RsFnumber("sNumber").Value
               'New getlastnumer as requested by lily ho baho
               
               Tmp = RsFnumber("sNumber").Value
               myTmpPos = Int(ReturnFirst(Tmp)) - (Int(ReturnFirst(Tmp)) - Int(Return_1stDash(Tmp)))
               Tmp = Mid(Tmp, Return_1stDash(Tmp) + 5, myTmpPos - 1)
               GetLastNumber = "SD" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & AutoIncrement(Tmp)
               
        Else
               GetLastNumber = "SD" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & "000000000"
        End If
End With

End Function



Public Function returnMon() As String
Dim Tmp
If Len(Month(Now)) <= 1 Then
Tmp = "0" & Month(Now)
Else
Tmp = Month(Now)
End If
returnMon = Tmp
End Function

Public Function returnDay() As String
Dim Tmp
If Len(Day(Now)) <= 1 Then
Tmp = "0" & Day(Now)
Else
Tmp = Day(Now)
End If
returnDay = Tmp
End Function




Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub




Private Sub txtAccNo_LostFocus()
txtAccNo = UCase(txtAccNo)
End Sub

Private Sub txtAgencyName_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtCommPercent_LostFocus()
txtCommPercent.Enabled = False
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtDepDate1_Change(Index As Integer)
If Index = 0 Then
    For i = 1 To 14
        Me.txtDepDate1(i).Text = Me.txtDepDate1(Index).Text
    Next i
End If
End Sub

Private Sub txtDepDate1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtDepDate2_Change(Index As Integer)
If Index = 0 Then
    For i = 1 To 14
        Me.txtDepDate2(i).Text = Me.txtDepDate2(Index).Text
    Next i
End If
End Sub

Private Sub txtDepDate2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtDepDate3_Change(Index As Integer)
If Index = 0 Then
    For i = 1 To 14
        Me.txtDepDate3(i).Text = Me.txtDepDate3(Index).Text
    Next i
End If
End Sub

Function GetRouteNoShowFee(Param, Index, Optional PriceSelector) As Double
Dim i                           As Integer
Dim Tmp(1 To 2)                 As String
Dim Rstmp                       As ADODB.Recordset
Dim RsAmt                       As ADODB.Recordset

'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

'=========================================
'my mod
'=========================================
Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)

'=========================================


Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"

With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
If .RecordCount > 0 Then

SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
    .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
    " AND [FareBasis]='" & PriceSelector & "'"
    
    
    RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
                If RsAmt.RecordCount > 0 Then
                
                            GetRouteNoShowFee = RsAmt.Fields("Noshow Fee").Value
                        Else
                            GetRouteNoShowFee = 0
                End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function


Function GetRouteVoidFee(Param, Index, Optional PriceSelector) As Double
Dim i                           As Integer
Dim Tmp(1 To 2)                 As String
Dim Rstmp                       As ADODB.Recordset
Dim RsAmt                       As ADODB.Recordset

'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

'=========================================
'my mod
'=========================================
Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)

'=========================================


Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"

With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
If .RecordCount > 0 Then

SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
    .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
    " AND [FareBasis]='" & PriceSelector & "'"
    
    
    RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
                If RsAmt.RecordCount > 0 Then
                
                            GetRouteVoidFee = RsAmt.Fields("Void Fee").Value
                        Else
                            GetRouteVoidFee = 0
                End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function


Function GetRouteRefundFee(Param, Index, Optional PriceSelector) As Double
Dim i                           As Integer
Dim Tmp(1 To 2)                 As String
Dim Rstmp                       As ADODB.Recordset
Dim RsAmt                       As ADODB.Recordset

'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

'=========================================
'my mod
'=========================================
Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)

'=========================================


Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"

With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
If .RecordCount > 0 Then

SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
    .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
    " AND [FareBasis]='" & PriceSelector & "'"
    
    
    RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
                If RsAmt.RecordCount > 0 Then
                
                            GetRouteRefundFee = RsAmt.Fields("Refund Fee").Value
                        Else
                            GetRouteRefundFee = 0
                End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function

Function GetRouteServiceFee(Param, Index, Optional PriceSelector) As Double
Dim i                           As Integer
Dim Tmp(1 To 2)                 As String
Dim Rstmp                       As ADODB.Recordset
Dim RsAmt                       As ADODB.Recordset

'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

'=========================================
'my mod
'=========================================
Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)

'=========================================


Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"

With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
If .RecordCount > 0 Then

SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
    .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
    " AND [FareBasis]='" & PriceSelector & "'"
    
    
    RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
                If RsAmt.RecordCount > 0 Then
                
                            GetRouteServiceFee = RsAmt.Fields("Service Fee").Value
                        Else
                            GetRouteServiceFee = 0
                End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function


Function GetRouteEVAT(Param, Index, Optional PriceSelector) As Double
Dim i                           As Integer
Dim Tmp(1 To 2)                 As String
Dim Rstmp                       As ADODB.Recordset
Dim RsAmt                       As ADODB.Recordset

'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

'=========================================
'my mod
'=========================================
Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)

'=========================================


Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"

With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
If .RecordCount > 0 Then

SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
    .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
    " AND [FareBasis]='" & PriceSelector & "'"
    

    RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
                If RsAmt.RecordCount > 0 Then
                
                            GetRouteEVAT = RsAmt.Fields("EVAT").Value
                        Else
                            GetRouteEVAT = 0
                End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function


Function GetRouteAmt(Param, Index, Optional PriceSelector) As Double
Dim i                           As Integer
Dim Tmp(1 To 2)                 As String
Dim Rstmp                       As ADODB.Recordset
Dim RsAmt                       As ADODB.Recordset

'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

'=========================================
'my mod
'=========================================
Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)


'=========================================

Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"

With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
If .RecordCount > 0 Then

SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
    .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
    " AND [FareBasis]='" & PriceSelector & "'"
    
    
    RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
                If RsAmt.RecordCount > 0 Then
                
                            GetRouteAmt = RsAmt.Fields("Gross Fare").Value
                        Else
                            GetRouteAmt = 0
                End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function

Sub Recalc(Index)
On Error Resume Next
Dim MyTmpGross                      As Double
Dim MyTempNet                       As Double

Dim MyTempCboDest1                  As Double
Dim MyTempCboDest2                  As Double
Dim MyTempCboFrom                   As Double

Dim MyTempNetAmountFrom             As Double
Dim MyTempNetAmountDest1            As Double
Dim MyTempNetAmountDest2            As Double
Dim MyTempNetAmountDest3            As Double

Dim MyTempCompare(1 To 4)           As String

Dim SelectPriceFrom(1 To 3)         As String
Dim SelectPriceDest1(1 To 3)        As String
Dim SelectPriceDest2(1 To 3)        As String
Dim SelectPriceDest3(1 To 3)        As String




    txtGrossFare(Index) = "0.00"


'============================================================
'my old code
'MyTempCompare(1) = Mid$(Me.cboFrom(Index), 8, 3)
'MyTempCompare(2) = Mid$(Me.cboDest1(Index), 8, 3)
'MyTempCompare(3) = Mid$(Me.cboDest2(Index), 8, 3)
'MyTempCompare(4) = Mid$(Me.cboDest3(Index), 8, 3)
'============================================================

MyTempCompare(1) = Mid$(Me.cboFrom(Index), ReturnPos(Me.cboFrom(Index)) + 1, Len(Me.cboFrom(Index)))
MyTempCompare(2) = Mid$(Me.cboDest1(Index), ReturnPos(Me.cboDest1(Index)) + 1, Len(Me.cboDest1(Index)))
MyTempCompare(3) = Mid$(Me.cboDest2(Index), ReturnPos(Me.cboDest2(Index)) + 1, Len(Me.cboDest2(Index)))
MyTempCompare(4) = Mid$(Me.cboDest3(Index), ReturnPos(Me.cboDest3(Index)) + 1, Len(Me.cboDest3(Index)))


myRouteEVAT1(Index) = GetRouteEVAT(Me.cboFrom(Index), Index, MyTempCompare(1))
myRouteEVAT2(Index) = GetRouteEVAT(Me.cboDest1(Index), Index, MyTempCompare(2))
myRouteEVAT3(Index) = GetRouteEVAT(Me.cboDest2(Index), Index, MyTempCompare(3))
myRouteEVAT4(Index) = GetRouteEVAT(Me.cboDest3(Index), Index, MyTempCompare(4))

myRouteServiceFEE1(Index) = GetRouteServiceFee(Me.cboFrom(Index), Index, MyTempCompare(1))
myRouteServiceFEE2(Index) = GetRouteServiceFee(Me.cboDest1(Index), Index, MyTempCompare(2))
myRouteServiceFEE3(Index) = GetRouteServiceFee(Me.cboDest2(Index), Index, MyTempCompare(3))
myRouteServiceFEE4(Index) = GetRouteServiceFee(Me.cboDest3(Index), Index, MyTempCompare(4))

myRouteRefundFEE1(Index) = GetRouteRefundFee(Me.cboFrom(Index), Index, MyTempCompare(1))
myRouteRefundFEE2(Index) = GetRouteRefundFee(Me.cboDest1(Index), Index, MyTempCompare(2))
myRouteRefundFEE3(Index) = GetRouteRefundFee(Me.cboDest2(Index), Index, MyTempCompare(3))
myRouteRefundFEE4(Index) = GetRouteRefundFee(Me.cboDest3(Index), Index, MyTempCompare(4))

myRouteVoidFEE1(Index) = GetRouteVoidFee(Me.cboFrom(Index), Index, MyTempCompare(1))
myRouteVoidFEE2(Index) = GetRouteVoidFee(Me.cboDest1(Index), Index, MyTempCompare(2))
myRouteVoidFEE3(Index) = GetRouteVoidFee(Me.cboDest2(Index), Index, MyTempCompare(3))
myRouteVoidFEE4(Index) = GetRouteVoidFee(Me.cboDest3(Index), Index, MyTempCompare(4))

myRouteNoShowFEE1(Index) = GetRouteNoShowFee(Me.cboFrom(Index), Index, MyTempCompare(1))
myRouteNoShowFEE2(Index) = GetRouteNoShowFee(Me.cboDest1(Index), Index, MyTempCompare(2))
myRouteNoShowFEE3(Index) = GetRouteNoShowFee(Me.cboDest2(Index), Index, MyTempCompare(3))
myRouteNoShowFEE4(Index) = GetRouteNoShowFee(Me.cboDest3(Index), Index, MyTempCompare(4))


myArr_Evat(Index) = CDbl(myRouteEVAT1(Index)) + CDbl(myRouteEVAT2(Index)) + CDbl(myRouteEVAT3(Index)) + CDbl(myRouteEVAT4(Index))
myArr_ServiceFee(Index) = CDbl(myRouteServiceFEE1(Index)) + CDbl(myRouteServiceFEE2(Index)) + CDbl(myRouteServiceFEE3(Index)) + CDbl(myRouteServiceFEE4(Index))
myArr_RefundFee(Index) = CDbl(myRouteRefundFEE1(Index)) + CDbl(myRouteRefundFEE2(Index)) + CDbl(myRouteRefundFEE3(Index)) + CDbl(myRouteRefundFEE4(Index))

myArr_VoidFee(Index) = CDbl(myRouteVoidFEE1(Index)) + CDbl(myRouteVoidFEE2(Index)) + CDbl(myRouteVoidFEE3(Index)) + CDbl(myRouteVoidFEE4(Index))
myArr_NoShowFee(Index) = CDbl(myRouteNoShowFEE1(Index)) + CDbl(myRouteNoShowFEE2(Index)) + CDbl(myRouteNoShowFEE3(Index)) + CDbl(myRouteNoShowFEE4(Index))


MyTempCboFrom = CDbl(GetRouteAmt(Me.cboFrom(Index), Index, MyTempCompare(1)))
MyTempCboDest1 = CDbl(GetRouteAmt(Me.cboDest1(Index), Index, MyTempCompare(2)))
MyTempCboDest2 = CDbl(GetRouteAmt(Me.cboDest2(Index), Index, MyTempCompare(3)))
MyTempCboDest3 = CDbl(GetRouteAmt(Me.cboDest3(Index), Index, MyTempCompare(4)))


    MyTmpGross = MyTempCboFrom + MyTempCboDest1 + MyTempCboDest2 + MyTempCboDest3

MyTempNetAmountFrom = CDbl(GetRouteNetAmt(Me.cboFrom(Index), Index, 1, MyTempCompare(1)))
MyTempNetAmountDest1 = CDbl(GetRouteNetAmt(Me.cboDest1(Index), Index, 2, MyTempCompare(2)))
MyTempNetAmountDest2 = CDbl(GetRouteNetAmt(Me.cboDest2(Index), Index, 3, MyTempCompare(3)))
MyTempNetAmountDest3 = CDbl(GetRouteNetAmt(Me.cboDest3(Index), Index, 4, MyTempCompare(4)))


    MyTempNet = MyTempNetAmountFrom + MyTempNetAmountDest1 + MyTempNetAmountDest2 + MyTempNetAmountDest3


'SD106-011801311-2



If Me.cboPassengerType(Index).Text <> "VOID" Then

    txtGrossFare(Index) = Format(CDbl(MyTmpGross), "###,##0.00")
    txtNet(Index) = Format(CDbl(MyTempNet), "###,##0.00")


Me.txtFare(Index) = Format(CDbl(txtNet(Index)) - (CDbl(txtNet(Index)) * (CDbl(txtDiscountAmt(Index)) / 100)) _
                  - (CDbl(txtGrossFare(Index)) * (CDbl(txtComm(Index)) / 100)), "###,##0.00")
End If

End Sub

Function ReturnRoute(Param, flag) As String
Dim myCtr As Integer
Dim myPointer(1 To 2) As Integer

'=========================================
'my mod
'=========================================
myCtr = 0
myPointer(1) = 0
myPointer(2) = 0

For i = 1 To Len(Param) Step 1
    If Mid(Param, i, 1) = "/" Then
        myCtr = myCtr + 1
        myPointer(myCtr) = i
    End If
Next i
If flag = 1 Then
    ReturnRoute = Mid$(Param, 1, myPointer(1) - 1)
Else
    ReturnRoute = Mid$(Param, myPointer(1) + 1, myPointer(2) - (myPointer(1) + 1))
End If

End Function


Function ReturnPos(Param) As Integer
Dim myCtr As Integer
Dim myPointer(1 To 2) As Integer
myCtr = 0
myPointer(1) = 0
For i = 1 To Len(Param) Step 1
    If Mid(Param, i, 1) = "/" Then
        myCtr = myCtr + 1
        
        If myCtr = 2 Then
                myPointer(1) = i
        End If
    End If
Next i
    ReturnPos = myPointer(1)
    
End Function

Function GetRouteNetAmt(Param, Index, RouteCtr, Optional PriceSelector) As Double
On Error Resume Next
Dim i As Integer
Static Tmp(1 To 2) As String
Dim Rstmp As ADODB.Recordset
Dim RsAmt As ADODB.Recordset
Dim RstCloneTmp As New ADODB.Recordset

'=========================================
'my old code
'Tmp(1) = Mid$(Param, 1, 3)
'Tmp(2) = Mid$(Param, 4, 3)

Tmp(1) = ReturnRoute(Param, 1)
Tmp(2) = ReturnRoute(Param, 2)

'=========================================



Set Rstmp = New ADODB.Recordset
Set RsAmt = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Routes WHERE [FROM]='" & Tmp(1) & "' AND [TO]='" & Tmp(2) & "'"
With Rstmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        
            SQL = "SELECT * FROM tbl_RoutePricing WHERE [RouteID]=" & _
                .Fields(0).Value & " AND [AirlineID]=" & FindAirline(Me.Combo1) & _
                " AND [FareBasis]='" & PriceSelector & "'"
                
            
         
                RsAmt.Open SQL, cn, adOpenKeyset, adLockOptimistic
    
    
                If RsAmt.RecordCount > 0 Then
                'GetRouteNetAmt = RsAmt.Fields("Net Fare").Value
                
                Dim MyTempGross As Double
                Dim MyTempInsurance As Double
                Dim MyTempASF As Double
                Dim MyTempTF As Double
                Dim MyTempMeals As Double
                Dim MyTempMisc As Double
                                
                MyTempGross = IIf(IsNull(RsAmt.Fields("Gross Fare").Value), 0, RsAmt.Fields("Gross Fare").Value)
                MyTempInsurance = IIf(IsNull(RsAmt.Fields("Insurance").Value), 0, RsAmt.Fields("Insurance").Value)
                MyTempASF = IIf(IsNull(RsAmt.Fields("ASF").Value), 0, RsAmt.Fields("ASF").Value)
                MyTempTF = IIf(IsNull(RsAmt.Fields("Terminal Fee").Value), 0, RsAmt.Fields("Terminal Fee").Value)
                MyTempMeals = IIf(IsNull(RsAmt.Fields("Meals").Value), 0, RsAmt.Fields("Meals").Value)
                MyTempMisc = IIf(IsNull(RsAmt.Fields("Misc").Value), 0, RsAmt.Fields("Misc").Value)
                                
                If UCase(PriceSelector) = RsAmt.Fields("FareBasis").Value Then
                        'GetRouteNetAmt = MyTempGross + MyTempInsurance + MyTempASF + MyTempTF + MyTempMeals
                        GetRouteNetAmt = IIf(IsNull(RsAmt.Fields("Net Fare").Value), 0, RsAmt.Fields("Net Fare").Value)
                End If
                            CompanyComm(1) = IIf(IsNull(RsAmt.Fields("Commision").Value), 0, RsAmt.Fields("Commision").Value)
                            Insurance(1) = IIf(IsNull(RsAmt.Fields("Insurance").Value), 0, RsAmt.Fields("Insurance").Value)
                            ASF(1) = IIf(IsNull(RsAmt.Fields("ASF").Value), 0, RsAmt.Fields("ASF").Value)
                            TerminalFee(1) = IIf(IsNull(RsAmt.Fields("Terminal Fee").Value), 0, RsAmt.Fields("Terminal Fee").Value)
                            Meals(1) = IIf(IsNull(RsAmt.Fields("Meals").Value), 0, RsAmt.Fields("Meals").Value)
                            Misc(1) = IIf(IsNull(RsAmt.Fields("Misc").Value), 0, RsAmt.Fields("Misc").Value)
                                            For i = 0 To 14 Step 1
                                            If RouteCtr = 1 Then
                                                If Me.cboFrom(i).Text <> "" Then
                                                        txtCompaComm1(i).Text = CompanyComm(1)
                                                        Me.txtInsurance1(i).Text = Insurance(1)
                                                        Me.txtASF1(i).Text = ASF(1)
                                                        Me.txtTerminalFee1(i).Text = TerminalFee(1)
                                                        Me.txtMeals1(i).Text = Meals(1)
                                                        Me.txtMisc_R1(i).Text = Misc(1)
                                                End If
                                            End If
                                            
                                            If RouteCtr = 2 Then
                                                If Me.cboDest1(i).Text <> "" Then
                                                    txtCompaComm2(i).Text = CompanyComm(1)
                                                    Me.txtInsurance2(i).Text = Insurance(1)
                                                    Me.txtASF2(i).Text = ASF(1)
                                                    Me.txtTerminalFee2(i).Text = TerminalFee(1)
                                                    Me.txtMeals2(i).Text = Meals(1)
                                                    Me.txtMisc_R2(i).Text = Misc(1)
                                                End If
                                            End If
                                            
                                            
                                            If RouteCtr = 3 Then
                                                If Me.cboDest2(i).Text <> "" Then
                                                    txtCompaComm3(i).Text = CompanyComm(1)
                                                    Me.txtInsurance3(i).Text = Insurance(1)
                                                    Me.txtASF3(i).Text = ASF(1)
                                                    Me.txtTerminalFee3(i).Text = TerminalFee(1)
                                                    Me.txtMeals3(i).Text = Meals(1)
                                                    Me.txtMisc_R3(i).Text = Misc(1)
                                                End If
                                            End If
                                            
                                            
                                            If RouteCtr = 4 Then
                                                If Me.cboDest3(i).Text <> "" Then
                                                    txtCompaComm4(i).Text = CompanyComm(1)
                                                    Me.txtInsurance4(i).Text = Insurance(1)
                                                    Me.txtASF4(i).Text = ASF(1)
                                                    Me.txtTerminalFee4(i).Text = TerminalFee(1)
                                                    Me.txtMeals4(i).Text = Meals(1)
                                                    Me.txtMisc_R4(i).Text = Misc(1)
                                                End If
                                            
                                            End If

                                            DoEvents
                                            Next i
                        Else
                        'MsgBox "Please Select first Ticket Type"
                        'Call Clear
                            GetRouteNetAmt = 0
                            CompanyComm(1) = 0
                            Insurance(1) = 0
                            ASF(1) = 0
                            TerminalFee(1) = 0
                            Meals(1) = 0
               End If
                RsAmt.Close
                
                Set RsAmt = Nothing
        End If
      .Close
Set Rstmp = Nothing
End With
End Function

Function Check_Update_Ticket(Param, Optional flag As String) As Boolean
Dim RsCheck As ADODB.Recordset
Set RsCheck = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Tickets WHERE [Ticket No]='" & Param & "' AND [AirlineID]=" & FindAirline(Me.Combo1) & " AND [Status]='Un-Sold'"
With RsCheck
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            
            If flag = "Search" Or flag = "" Then
                    If .RecordCount > 0 Then
                            Check_Update_Ticket = True
                        Else
                          If Me.Tag = "over_ride" Then
                            Check_Update_Ticket = True
                          Else
                            Check_Update_Ticket = False
                          End If
                    End If
            Else
            If flag = "Edit" Then
            
                    If .RecordCount > 0 Then
                    On Error GoTo FailSafe_Err
                    cn.BeginTrans
                        .Fields("Status").Value = "Sold"
                        .Fields("Issued By").Value = Me.txtIssuedBy
                        .Fields("Date Issued").Value = Format(Me.txtDate, "mm/dd/yyyy")
                        .Fields("Statement No").Value = Me.txtNo
                        .Update
                    cn.CommitTrans
                    End If
            End If
            End If
           .Close
End With
Exit Function
FailSafe_Err:
cn.RollbackTrans
End Function

Function FindTicketTypeID(ByVal Param As String) As Long
Dim RsLook As New ADODB.Recordset
Dim STRSQL As String
SQL = "SELECT * FROM tbl_Tickets WHERE [Ticket No]='" & Param & "'"
With RsLook
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FindTicketTypeID = .Fields(3).Value
            Else
            FindTicketTypeID = -1
        End If
        .Close
End With
Set RsLook = Nothing
End Function

Function ReturnTicketType(ByVal Param As Long) As String
Dim RsLook As New ADODB.Recordset
Dim STRSQL As String
SQL = "SELECT * FROM tbl_TicketType WHERE [TicketTypeID]=" & Param
With RsLook
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            ReturnTicketType = .Fields(1).Value
            Else
            ReturnTicketType = ""
        End If
        .Close
End With
Set RsLook = Nothing

End Function

Private Sub txtDepDate3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub


Private Sub txtPassengerName_GotFocus(Index As Integer)
   If Check_Update_Ticket(txtTicketNo(Index), "Search") Then
            Me.cboTicketType(Index) = ReturnTicketType(FindTicketTypeID(Me.txtTicketNo(Index)))
         If CheckNull(Me.txtPassengerName) Then
            Me.cboFrom(Index).Clear
            Me.cboDest1(Index).Clear
            Me.cboDest2(Index).Clear
            Me.cboDest3(Index).Clear
            FillRoutes (FindTicketTypeID(Me.txtTicketNo(Index)))
         End If
    Else
        If Me.txtTicketNo(Index) <> "" Then
           MsgBox "This ticket number did not exist or already sold!" & Chr(13) & _
               "HINT!: Check the Airline/Shipping line if they have such ticket..", vbCritical
            txtTicketNo(Index).SelStart = 0
            txtTicketNo(Index).SelLength = Len(txtTicketNo(Index).Text)
            txtTicketNo(Index).SetFocus
       End If
    End If
End Sub

Private Sub txtPassengerName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtTicketNo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If CheckNull(txtTicketNo(Index)) Then: Exit Sub
If KeyCode = 13 Then
    If Check_Update_Ticket(txtTicketNo(Index), "Search") Then
            Me.cboTicketType(Index) = ReturnTicketType(FindTicketTypeID(Me.txtTicketNo(Index)))
            
            Me.cboFrom(Index).Clear
            Me.cboDest1(Index).Clear
            Me.cboDest2(Index).Clear
            Me.cboDest3(Index).Clear
            
            FillRoutes (FindTicketTypeID(Me.txtTicketNo(Index)))
            TrapEnter KeyCode
    Else
            If Me.txtTicketNo(Index) <> "" Then
                MsgBox "This ticket number did not exist or already sold!" & Chr(13) & _
                       "HINT!: Check the Airline/Shipping line if they have such ticket..", vbCritical
                txtTicketNo(Index).SelStart = 0
                txtTicketNo(Index).SelLength = Len(txtTicketNo(Index).Text)
                txtTicketNo(Index).SetFocus
            End If
    End If
End If

End Sub


Function ReturnAccDetails(ByVal UserAirlineID As Long, ByVal UserAccountId As Long, ByVal UserFld As String) As Double
Dim Rstmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_custAccDetails WHERE [AirlineID]=" & CLng(UserAirlineID) & " AND [AccountID]=" & CLng(UserAccountId)
With Rstmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                ReturnAccDetails = .Fields(UserFld).Value
            Else
                ReturnAccDetails = 0
            End If
           .Close
        Set Rstmp = Nothing
End With
End Function
