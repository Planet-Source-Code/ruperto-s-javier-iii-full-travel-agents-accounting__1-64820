VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmPassporting 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Passporting"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   Icon            =   "frmPassporting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10605
      TabIndex        =   45
      Top             =   -30
      Width           =   10665
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmPassporting.frx":6852
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Passporting and Customers documentation"
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
         TabIndex        =   47
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Passporting and Documentation"
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
         TabIndex        =   46
         Top             =   120
         Width           =   7695
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   7800
      Top             =   840
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
            Picture         =   "frmPassporting.frx":711C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":7DF6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":8C48
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":9A9A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":A374
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":AC4E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":B528
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":BEF2
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":C7CC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":CAE6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":D3C0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":DC9A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":E574
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":E88E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":F168
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":FA42
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":1031C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":10BF6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":114D0
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":11DAA
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":12684
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":12F5E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":13838
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":14112
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":149EC
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":152C6
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":15BA0
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":1647A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":16D54
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":1762E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":17EE4
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":187BE
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":18C10
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":19062
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPassporting.frx":1B814
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   60
      ScaleHeight     =   5025
      ScaleWidth      =   10470
      TabIndex        =   27
      Top             =   2805
      Width           =   10500
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8265
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   4410
         Width           =   2055
      End
      Begin VB.TextBox txtOthersDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   14
         Top             =   3165
         Width           =   3405
      End
      Begin VB.TextBox txtSECPADesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   12
         Top             =   2760
         Width           =   3405
      End
      Begin VB.TextBox txtReconfirmationDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   10
         Top             =   2385
         Width           =   3405
      End
      Begin VB.TextBox txtVisaDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   8
         Top             =   1965
         Width           =   3405
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8265
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   3105
         Width           =   2070
      End
      Begin VB.TextBox txtSECPA 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8265
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2745
         Width           =   2070
      End
      Begin VB.TextBox txtReconfirmation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8265
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   2355
         Width           =   2070
      End
      Begin VB.TextBox txtVisa 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8265
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1965
         Width           =   2070
      End
      Begin VB.TextBox txtExpress 
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
         Height          =   360
         Left            =   8280
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   780
         Width           =   2040
      End
      Begin VB.TextBox txtPassportAmount 
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
         Height          =   360
         Left            =   8265
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   375
         Width           =   2055
      End
      Begin LVbuttons.LaVolpeButton cmdNew 
         Height          =   480
         Left            =   105
         TabIndex        =   0
         Top             =   4425
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
         MICON           =   "frmPassporting.frx":1CA96
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
         Left            =   1470
         TabIndex        =   43
         Top             =   4425
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
         MICON           =   "frmPassporting.frx":1CAB2
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
         Left            =   2835
         TabIndex        =   16
         Top             =   4425
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
         MICON           =   "frmPassporting.frx":1CACE
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
         Left            =   4740
         TabIndex        =   44
         Top             =   4425
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
         MICON           =   "frmPassporting.frx":1CAEA
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "======================="
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   8250
         TabIndex        =   42
         Top             =   3900
         Width           =   2040
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount:"
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
         Index           =   18
         Left            =   6300
         TabIndex        =   40
         Top             =   4455
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DOCUMENTATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   255
         TabIndex        =   39
         Top             =   1515
         Width           =   7995
      End
      Begin VB.Shape Shape1 
         Height          =   2190
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1605
         Width           =   8010
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "===================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   6270
         TabIndex        =   38
         Top             =   2010
         Width           =   1845
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "==========================================================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2775
         TabIndex        =   37
         Top             =   855
         Width           =   5970
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "==========================================================>"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2790
         TabIndex        =   36
         Top             =   465
         Width           =   5970
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   480
         TabIndex        =   35
         Top             =   3255
         Width           =   750
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECPA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   480
         TabIndex        =   34
         Top             =   2880
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reconfirmation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   495
         TabIndex        =   33
         Top             =   2505
         Width           =   1635
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   525
         TabIndex        =   32
         Top             =   2085
         Width           =   540
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
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   8280
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
         Left            =   8265
         TabIndex        =   30
         Top             =   0
         Width           =   2070
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Express:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   510
         TabIndex        =   29
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passport Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   495
         TabIndex        =   28
         Top             =   450
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   60
      ScaleHeight     =   1365
      ScaleWidth      =   10470
      TabIndex        =   19
      Top             =   1380
      Width           =   10500
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   8655
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   45
         Width           =   1680
      End
      Begin VB.TextBox txtPassportNo 
         Height          =   285
         Left            =   7470
         TabIndex        =   5
         Top             =   960
         Width           =   2880
      End
      Begin VB.TextBox txtTelno 
         Height          =   285
         Left            =   1605
         TabIndex        =   3
         Top             =   930
         Width           =   2025
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1605
         TabIndex        =   2
         Top             =   510
         Width           =   3780
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1605
         TabIndex        =   1
         Top             =   120
         Width           =   3810
      End
      Begin MSComCtl2.DTPicker txtExpiryDate 
         Height          =   315
         Left            =   8655
         TabIndex        =   4
         Top             =   420
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53608449
         CurrentDate     =   38518
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   7470
         TabIndex        =   26
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   7470
         TabIndex        =   24
         Top             =   75
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passport No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6015
         TabIndex        =   23
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   975
         Width           =   570
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   255
         TabIndex        =   21
         Top             =   555
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   270
         TabIndex        =   20
         Top             =   165
         Width           =   1155
      End
   End
   Begin VB.TextBox txtStatementNo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   975
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statement No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   17
      Top             =   990
      Width           =   1470
   End
End
Attribute VB_Name = "frmPassporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
txtStatementNo = AutoIncrement(GetLastNumber)
Me.txtName.SetFocus
End Sub

Private Sub cmdPost_Click()
On Error GoTo ErrExit
Dim Rs As ADODB.Recordset
Dim SQL As String

ask = MsgBox("Are you sure you want to save this?", vbInformation + vbYesNo)
If ask = vbYes Then
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_PPT"
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        cn.BeginTrans
        .AddNew
        .Fields(1).Value = Me.txtStatementNo
        .Fields(2).Value = Me.txtPassportNo
        .Fields(3).Value = Me.txtName
        .Fields(4).Value = Me.txtAddress
        .Fields(5).Value = Me.txtTelNo
        .Fields(6).Value = Me.txtDate
        .Fields(7).Value = Me.txtExpiryDate
        .Fields(8).Value = Me.txtPassportAmount
        .Fields(9).Value = Me.txtExpress
        .Fields(10).Value = Me.txtVisa
        .Fields(11).Value = Me.txtReconfirmation
        .Fields(12).Value = Me.txtSECPA
        .Fields(13).Value = Me.txtOthers
        .Fields(14).Value = Me.txtVisaDesc
        .Fields(15).Value = Me.txtReconfirmationDesc
        .Fields(16).Value = Me.txtSECPADesc
        .Fields(17).Value = Me.txtOthersDesc
        .Fields(18).Value = Me.txtTotalAmount
        .Fields(19).Value = WhichBranch.Fields(2).Value
        
           .Fields("Paid").Value = False
           .Fields("Void").Value = False
           .Fields("Refund").Value = False
           .Fields("Balance").Value = CDbl(Me.txtTotalAmount)
           .Fields("Credit Card Activated").Value = False
        
        .Update
        cn.CommitTrans
        MsgBox "Record Save...", vbInformation
End With
End If
Exit Sub
ErrExit:
cn.RollbackTrans
End Sub

Private Sub Form_Load()
txtDate = Format(Now, "MM/DD/YYYY")
End Sub


Function GetLastNumber() As String
Dim RsFnumber As ADODB.Recordset
Dim SQL As String

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT StatementNo from tbl_PPT"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               GetLastNumber = RsFnumber("StatementNo").Value
        Else
               GetLastNumber = "SP" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnMon & "00000"
        End If
End With

End Function


Sub ReCompute()
txtTotalAmount = CDbl(txtPassportAmount) + CDbl(txtExpress) + CDbl(txtVisa) + _
                 CDbl(txtReconfirmation) + CDbl(txtSECPA) + CDbl(txtOthers)
txtTotalAmount = Format(txtTotalAmount, "###,##0.00")
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtExpiryDate_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtExpress_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtExpress) Then
    With txtExpress
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

Private Sub txtExpress_GotFocus()
txtExpress = Format(txtExpress, "###,##0.00")
With txtExpress
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
End With
End Sub

Private Sub txtExpress_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtExpress_LostFocus()
txtExpress = Format(txtExpress, "###,##0.00")
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtOthers_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtOthers) Then
    With txtOthers
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

Private Sub txtOthers_GotFocus()
txtOthers = Format(txtOthers, "###,##0.00")
With txtOthers
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
End With
End Sub

Private Sub txtOthers_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtOthers_LostFocus()
txtOthers = Format(txtOthers, "###,##0.00")
End Sub

Private Sub txtOthersDesc_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtPassportAmount_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtPassportAmount) Then
    With txtPassportAmount
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

Private Sub txtPassportAmount_GotFocus()
txtPassportAmount = Format(txtPassportAmount, "###,##0.00")
With txtPassportAmount
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
End With
End Sub

Private Sub txtPassportAmount_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtPassportAmount_LostFocus()
txtPassportAmount = Format(txtPassportAmount, "###,##0.00")
End Sub

Private Sub txtPassportNo_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtReconfirmation_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtReconfirmation) Then
    With txtReconfirmation
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

Private Sub txtReconfirmation_GotFocus()
txtReconfirmation = Format(txtReconfirmation, "###,##0.00")
With txtReconfirmation
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
End With
End Sub

Private Sub txtReconfirmation_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtReconfirmation_LostFocus()
txtReconfirmation = Format(txtReconfirmation, "###,##0.00")
End Sub

Private Sub txtReconfirmationDesc_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtSECPA_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtSECPA) Then
    With txtSECPA
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

Private Sub txtSECPA_GotFocus()
txtSECPA = Format(txtSECPA, "###,##0.00")
With txtSECPA
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
End With
End Sub

Private Sub txtSECPA_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtSECPA_LostFocus()
txtSECPA = Format(txtSECPA, "###,##0.00")
End Sub

Private Sub txtSECPADesc_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtTelno_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtVisa_Change()
On Error GoTo ErrExit
If Not IsNumeric(txtVisa) Then
    With txtVisa
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

Private Sub txtVisa_GotFocus()
txtVisa = Format(txtVisa, "###,##0.00")
With txtVisa
          .Text = "0.00"
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
End With
End Sub

Private Sub txtVisa_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtVisa_LostFocus()
txtVisa = Format(txtVisa, "###,##0.00")
End Sub

Private Sub txtVisaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub
