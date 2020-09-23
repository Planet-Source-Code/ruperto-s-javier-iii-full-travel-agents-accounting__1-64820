VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmAddPurchaseOrder 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   Icon            =   "frmAddPurchaseOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPOnumber 
      Height          =   300
      Left            =   5370
      TabIndex        =   57
      Top             =   165
      Width           =   3075
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   7665
      Top             =   750
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
            Picture         =   "frmAddPurchaseOrder.frx":0CCA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":19A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":27F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":3648
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":3F22
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":47FC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":50D6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":5AA0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":637A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":6694
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":6F6E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":7848
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":8122
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":843C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":8D16
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":95F0
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":9ECA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":A7A4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":B07E
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":B958
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":C232
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":CB0C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":D3E6
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":DCC0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":E59A
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":EE74
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":F74E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":10028
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":10902
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":111DC
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":11A92
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":1236C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":127BE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":12C10
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPurchaseOrder.frx":153C2
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   15
      TabIndex        =   51
      Top             =   7035
      Width           =   8445
      Begin LVbuttons.LaVolpeButton cmdAddSave 
         Height          =   480
         Left            =   105
         TabIndex        =   52
         Top             =   225
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
         MICON           =   "frmAddPurchaseOrder.frx":16644
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
         Left            =   6765
         TabIndex        =   53
         Top             =   210
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
         MICON           =   "frmAddPurchaseOrder.frx":16660
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
         TabIndex        =   54
         Top             =   225
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
         MICON           =   "frmAddPurchaseOrder.frx":1667C
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
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   120
      TabIndex        =   48
      Top             =   6300
      Width           =   8325
      Begin VB.TextBox txtCvno 
         Height          =   345
         Left            =   750
         TabIndex        =   50
         Top             =   195
         Width           =   1545
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CV NO :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   49
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   7
      Left            =   5475
      TabIndex        =   47
      Top             =   5835
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   6
      Left            =   5475
      TabIndex        =   46
      Top             =   5415
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   5
      Left            =   5475
      TabIndex        =   45
      Top             =   4965
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   4
      Left            =   5475
      TabIndex        =   44
      Top             =   4545
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   3
      Left            =   5475
      TabIndex        =   43
      Top             =   4155
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   2
      Left            =   5475
      TabIndex        =   42
      Top             =   3720
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   1
      Left            =   5490
      TabIndex        =   41
      Top             =   3315
      Width           =   2955
   End
   Begin VB.TextBox To 
      Height          =   285
      Index           =   0
      Left            =   5490
      TabIndex        =   40
      Top             =   2880
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   7
      Left            =   2385
      TabIndex        =   39
      Top             =   5850
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   6
      Left            =   2385
      TabIndex        =   38
      Top             =   5430
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   5
      Left            =   2385
      TabIndex        =   37
      Top             =   4980
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   4
      Left            =   2385
      TabIndex        =   36
      Top             =   4560
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   3
      Left            =   2385
      TabIndex        =   35
      Top             =   4170
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   2
      Left            =   2385
      TabIndex        =   34
      Top             =   3735
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   33
      Top             =   3330
      Width           =   2955
   End
   Begin VB.TextBox From 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   32
      Top             =   2895
      Width           =   2955
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TAT"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   1575
      TabIndex        =   31
      Top             =   5835
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TAT"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   1575
      TabIndex        =   30
      Top             =   5400
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "MCO"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   1575
      TabIndex        =   29
      Top             =   4980
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "MCO"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   1575
      TabIndex        =   28
      Top             =   4560
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RT"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   1575
      TabIndex        =   27
      Top             =   4140
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RT"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   1575
      TabIndex        =   26
      Top             =   3750
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OW"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   1575
      TabIndex        =   25
      Top             =   3315
      Width           =   825
   End
   Begin VB.CheckBox chkType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OW"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   24
      Top             =   2865
      Width           =   825
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   5865
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   5445
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   5010
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   4590
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   4185
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   3765
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3330
      Width           =   900
   End
   Begin VB.TextBox txtPcs 
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2895
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1470
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8325
      Begin VB.TextBox txtAddress 
         Height          =   300
         Left            =   615
         TabIndex        =   6
         Top             =   975
         Width           =   6885
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Left            =   4440
         TabIndex        =   4
         Top             =   285
         Width           =   3060
      End
      Begin VB.ComboBox combo1 
         Height          =   315
         Left            =   615
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   3105
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ADDRESS :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   705
         Width           =   915
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DATE :"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TO:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PO NUMBER"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4200
      TabIndex        =   58
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TO"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5505
      TabIndex        =   56
      Top             =   2565
      Width           =   2955
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FROM"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2400
      TabIndex        =   55
      Top             =   2580
      Width           =   2955
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1095
      TabIndex        =   23
      Top             =   5865
      Width           =   420
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   22
      Top             =   5460
      Width           =   420
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1095
      TabIndex        =   21
      Top             =   5040
      Width           =   420
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1095
      TabIndex        =   20
      Top             =   4605
      Width           =   420
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   3780
      Width           =   420
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1095
      TabIndex        =   17
      Top             =   3345
      Width           =   420
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PCS."
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Top             =   2895
      Width           =   420
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "PURCHASE ORDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   1965
      Width           =   8340
   End
End
Attribute VB_Name = "frmAddPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddSave_Click()
Dim ask As Integer
Dim SQL As String
Dim Rs As New ADODB.Recordset

SQL = "SELECT * FROM tbl_PurchaseOrder"

ask = MsgBox("Are you sure you want to save?", vbInformation + vbYesNo)
If ask = vbYes Then
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("PoNumber").Value = Me.txtPOnumber
        .Fields("AirlineID").Value = FindAirline(Me.combo1)
        .Fields("Date").Value = Format(Me.txtDate, "mm/dd/yyyy")
        .Fields("Address").Value = Me.txtAddress
        .Fields("CVno").Value = Me.txtCvno
        .Update
End With
End If
End Sub

Function FindAirline(param) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & param & "'"
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Sub FillCombo()
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline"
Me.combo1.Clear
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
        Me.combo1.AddItem .Fields(1).Value
        .MoveNext
        Loop
    End If
End With
Me.combo1.ListIndex = 0

End Sub

Private Sub Form_Load()
Me.txtDate = Format(Now, "mm/dd/yyyy")
Call FillCombo
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
        End If
End With

End Function
