VERSION 5.00
Object = "{CA9A4B18-8725-45C0-A7C2-226B0CE1A0D0}#4.0#0"; "imgXPTabs.ocx"
Begin VB.Form frmCashierPaymentOpt 
   BorderStyle     =   0  'None
   ClientHeight    =   5145
   ClientLeft      =   7470
   ClientTop       =   3840
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin Accounting.ucGradContainer ucGradContainer1 
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   9075
      IconSize        =   0
      BackColor1      =   16577775
      Caption         =   "Select Which Type of Payment"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   450
         Left            =   4530
         TabIndex        =   2
         Top             =   4500
         Width           =   1485
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   450
         Left            =   2985
         TabIndex        =   3
         Top             =   4500
         Width           =   1485
      End
      Begin imgXPTabs.imgXpTab imgXpTab1 
         Height          =   3945
         Left            =   105
         TabIndex        =   1
         Top             =   465
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6959
         TabCount        =   4
         TabCaption(0)   =   "Cash"
         TabPicture(0)   =   "frmCashierPaymentOpt.frx":0000
         TabContCtrlCnt(0)=   2
         Tab(0)ContCtrlCap(1)=   "txtCashAmount"
         Tab(0)ContCtrlCap(2)=   "Label1"
         TabCaption(1)   =   "Card"
         TabPicture(1)   =   "frmCashierPaymentOpt.frx":3142
         TabContCtrlCnt(1)=   19
         Tab(1)ContCtrlCap(1)=   "txtAddress"
         Tab(1)ContCtrlCap(2)=   "txtyear"
         Tab(1)ContCtrlCap(3)=   "txtMonth"
         Tab(1)ContCtrlCap(4)=   "txtccv2"
         Tab(1)ContCtrlCap(5)=   "txtCardHolder"
         Tab(1)ContCtrlCap(6)=   "txtCardNumber"
         Tab(1)ContCtrlCap(7)=   "txtCardAmount"
         Tab(1)ContCtrlCap(8)=   "CboCardName"
         Tab(1)ContCtrlCap(9)=   "txtMarkPercent"
         Tab(1)ContCtrlCap(10)=   "Label9"
         Tab(1)ContCtrlCap(11)=   "Label8"
         Tab(1)ContCtrlCap(12)=   "Label7"
         Tab(1)ContCtrlCap(13)=   "Label6"
         Tab(1)ContCtrlCap(14)=   "Label5"
         Tab(1)ContCtrlCap(15)=   "Label14"
         Tab(1)ContCtrlCap(16)=   "Label15"
         Tab(1)ContCtrlCap(17)=   "Label16"
         Tab(1)ContCtrlCap(18)=   "Label17"
         Tab(1)ContCtrlCap(19)=   "Label19"
         TabCaption(2)   =   "Check"
         TabPicture(2)   =   "frmCashierPaymentOpt.frx":432C
         TabContCtrlCnt(2)=   15
         Tab(2)ContCtrlCap(1)=   "txtNoofDays"
         Tab(2)ContCtrlCap(2)=   "OptNo"
         Tab(2)ContCtrlCap(3)=   "OptYes"
         Tab(2)ContCtrlCap(4)=   "txtCheckAmount"
         Tab(2)ContCtrlCap(5)=   "txtCheckNumber"
         Tab(2)ContCtrlCap(6)=   "txtCheckBank"
         Tab(2)ContCtrlCap(7)=   "txtCheckBranch"
         Tab(2)ContCtrlCap(8)=   "txtCheckDate"
         Tab(2)ContCtrlCap(9)=   "Label4"
         Tab(2)ContCtrlCap(10)=   "Label3"
         Tab(2)ContCtrlCap(11)=   "Label10"
         Tab(2)ContCtrlCap(12)=   "Label11"
         Tab(2)ContCtrlCap(13)=   "Label12"
         Tab(2)ContCtrlCap(14)=   "Label13"
         Tab(2)ContCtrlCap(15)=   "Label28"
         TabCaption(3)   =   "Others"
         TabPicture(3)   =   "frmCashierPaymentOpt.frx":24437E
         TabContCtrlCnt(3)=   2
         Tab(3)ContCtrlCap(1)=   "txtOthersAmount"
         Tab(3)ContCtrlCap(2)=   "Label2"
         TabTheme        =   1
         ActiveTabBackStartColor=   16514555
         ActiveTabBackEndColor=   16514555
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   15397104
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   10198161
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   10526880
         Begin VB.TextBox txtAddress 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72495
            TabIndex        =   40
            Top             =   2535
            Width           =   3150
         End
         Begin VB.TextBox txtyear 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -70365
            TabIndex        =   36
            Top             =   2940
            Width           =   990
         End
         Begin VB.TextBox txtMonth 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71835
            TabIndex        =   35
            Top             =   2940
            Width           =   765
         End
         Begin VB.TextBox txtccv2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72495
            TabIndex        =   33
            Top             =   2160
            Width           =   3150
         End
         Begin VB.TextBox txtNoofDays 
            Alignment       =   1  'Right Justify
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
            Left            =   -72600
            TabIndex        =   32
            Text            =   "0"
            Top             =   2820
            Width           =   3420
         End
         Begin VB.OptionButton OptNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -70680
            TabIndex        =   31
            Top             =   2505
            Width           =   1335
         End
         Begin VB.OptionButton OptYes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -72495
            TabIndex        =   30
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtOthersAmount 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   -72375
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   1560
            Width           =   2925
         End
         Begin VB.TextBox txtCardHolder 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72495
            TabIndex        =   20
            Top             =   1785
            Width           =   3150
         End
         Begin VB.TextBox txtCardNumber 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72495
            TabIndex        =   19
            Top             =   1410
            Width           =   3150
         End
         Begin VB.TextBox txtCardAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72495
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   645
            Width           =   1665
         End
         Begin VB.ComboBox CboCardName 
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
            ItemData        =   "frmCashierPaymentOpt.frx":247440
            Left            =   -72480
            List            =   "frmCashierPaymentOpt.frx":247456
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1035
            Width           =   3180
         End
         Begin VB.TextBox txtMarkPercent 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71565
            TabIndex        =   16
            Text            =   "6.00"
            Top             =   3375
            Width           =   2160
         End
         Begin VB.TextBox txtCheckAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   -72600
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   600
            Width           =   1230
         End
         Begin VB.TextBox txtCheckNumber 
            Alignment       =   1  'Right Justify
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
            Left            =   -72600
            TabIndex        =   9
            Top             =   975
            Width           =   3420
         End
         Begin VB.TextBox txtCheckBank 
            Alignment       =   1  'Right Justify
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
            Left            =   -72600
            TabIndex        =   8
            Top             =   1350
            Width           =   3420
         End
         Begin VB.TextBox txtCheckBranch 
            Alignment       =   1  'Right Justify
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
            Left            =   -72600
            TabIndex        =   7
            Top             =   1710
            Width           =   3420
         End
         Begin VB.TextBox txtCheckDate 
            Alignment       =   1  'Right Justify
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
            Left            =   -72600
            TabIndex        =   6
            Top             =   2085
            Width           =   3420
         End
         Begin VB.TextBox txtCashAmount 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2475
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   1290
            Width           =   2925
         End
         Begin VB.Label Label9 
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
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74775
            TabIndex        =   41
            Top             =   2565
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "YEAR"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -70905
            TabIndex        =   39
            Top             =   2955
            Width           =   435
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTH"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -72525
            TabIndex        =   38
            Top             =   2955
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CARD EXPIRE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74775
            TabIndex        =   37
            Top             =   2970
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CARD SECURITY CODE : "
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74775
            TabIndex        =   34
            Top             =   2235
            Width           =   2010
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO OF DAYS :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74865
            TabIndex        =   29
            Top             =   2910
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "POST DATED :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74880
            TabIndex        =   28
            Top             =   2535
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OTHERS AMOUNT :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -73965
            TabIndex        =   27
            Top             =   1545
            Width           =   1515
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CARD HOLDER :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74775
            TabIndex        =   25
            Top             =   1830
            Width           =   1305
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CARD NUMBER"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74790
            TabIndex        =   24
            Top             =   1455
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CARD NAME :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74790
            TabIndex        =   23
            Top             =   1065
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74775
            TabIndex        =   22
            Top             =   690
            Width           =   795
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MARK UP PERCENTAGE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74835
            TabIndex        =   21
            Top             =   3420
            Width           =   1995
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK AMOUNT :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74895
            TabIndex        =   15
            Top             =   645
            Width           =   1920
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK # :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74895
            TabIndex        =   14
            Top             =   1035
            Width           =   1275
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BANK NAME :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74895
            TabIndex        =   13
            Top             =   1395
            Width           =   1500
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BANK BRANCH :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74880
            TabIndex        =   12
            Top             =   1755
            Width           =   1530
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK DATE :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74880
            TabIndex        =   11
            Top             =   2190
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CASH AMOUNT :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   585
            TabIndex        =   5
            Top             =   1290
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmCashierPaymentOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL                 As String
Dim pIndex              As Long
Dim TmpSel              As Long
Dim myTemp_Accno        As String

Dim myCardFlag          As Boolean


Dim tmpBal              As Double
Dim tmpmarkup           As Double
Dim OldBal              As Double


Private Sub cmdExit_Click()
Unload Me
End Sub


'SD105-102500605-1
'SD105-102500007
Sub GroupBank_Account(ByVal Index As Long, ByVal pIndex As Long)
On Error GoTo FailSafe_Error
'//Assign Values

'//============================================================================================
'//   1.  PAL           -> PAL MBTC
'//============================================================================================
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_PAL Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_PAL_MBTC
        End If
'//============================================================================================
'//   2.  CP           -> CP MBTC
'//============================================================================================
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_CP Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_CP_MBTC
        End If
'//============================================================================================
'//   3.  CP           -> CK/NN/WGA/TA-EPCI
'//============================================================================================
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_CK Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_CK_NN_WGA_TA_EPCI
        End If
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_NN Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_CK_NN_WGA_TA_EPCI
        End If
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_WGA Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_CK_NN_WGA_TA_EPCI
        End If
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_TA Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_CK_NN_WGA_TA_EPCI
        End If

'//============================================================================================
'//   4.  CP           -> AP/AS-EPCI
'//============================================================================================
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_AP Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_AP_AS_EPCI
        End If
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_AS Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_AP_AS_EPCI
        End If
'//============================================================================================
'//   5.  Others           ->
'//============================================================================================
        If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = "" Then
                frmCashier.ListView2.ListItems(pIndex).SubItems(Index) = myGlobal_LilyHo_AccNo
        End If
        
        
Exit Sub
FailSafe_Error:
MsgBox "There was an error in returning the account #..."
End Sub

Private Sub cmdOk_Click()
pIndex = CDbl(Me.Tag)

Select Case TmpSel
'//-----------------------------------------------------------------------------------------------
'// For Cash
'//-----------------------------------------------------------------------------------------------

Case 0
        frmCashier.ListView2.ListItems(pIndex).SubItems(4) = Format(Me.txtCashAmount, "###,##0.00")
        If frmCashier.Option2 Then
            Call GroupBank_Account(20, pIndex)
        Else
            Call GroupBank_Account(20, pIndex)
        End If
        

'//-----------------------------------------------------------------------------------------------
'// For Credit Cards
'//-----------------------------------------------------------------------------------------------
Case 1

        frmCashier.ListView2.ListItems(pIndex).SubItems(6) = Format(Me.txtCardAmount, "###,##0.00")
        frmCashier.ListView2.ListItems(pIndex).SubItems(11) = Me.CboCardName
        frmCashier.ListView2.ListItems(pIndex).SubItems(12) = Me.txtCardHolder
        frmCashier.ListView2.ListItems(pIndex).SubItems(13) = Me.txtCardNumber
        frmCashier.ListView2.ListItems(pIndex).SubItems(14) = Format(Now, "mm/dd/yyyy")
        
        
        frmCashier.ListView2.ListItems(pIndex).SubItems(24) = Me.txtAddress
        frmCashier.ListView2.ListItems(pIndex).SubItems(25) = Me.txtMonth
        frmCashier.ListView2.ListItems(pIndex).SubItems(26) = Me.txtyear
        frmCashier.ListView2.ListItems(pIndex).SubItems(27) = Me.txtccv2
        
        If Me.CboCardName = "MASTERCARD" Or Me.CboCardName = "VISA" Then
           myTemp_Accno = Return_AccNo("MASTERCARD")
           Call GroupBank_Account(22, pIndex)
        End If
                
        If Me.CboCardName = "PAL-HSBC" Then
           myTemp_Accno = Return_AccNo("PAL HSBC")
                If frmCashier.ListView2.ListItems(pIndex).SubItems(2) = myGlobal_PAL Then
                        frmCashier.ListView2.ListItems(pIndex).SubItems(22) = myGlobal_PAL_HSBC
                End If
           
        End If
                        
        If Me.CboCardName = "DINERS" Then
           myTemp_Accno = Return_AccNo("DINERS")
           Call GroupBank_Account(22, pIndex)
        End If
        
        
        
        Dim tmpBal As Double
        Dim tmpmarkup As Double
        
        tmpBal = frmCashier.ListView2.ListItems(pIndex).SubItems(1)
        If CheckNull(Me.txtMarkPercent) Then
            Me.txtMarkPercent = "0.00"
        End If
        tmpmarkup = CDbl(tmpBal) * (CDbl(txtMarkPercent) / 100)
        
        frmCashier.ListView2.ListItems(pIndex).SubItems(1) = CDbl(tmpBal) + tmpmarkup
        frmCashier.ListView2.ListItems(pIndex).SubItems(19) = myTemp_Accno
        
        
'//-----------------------------------------------------------------------------------------------
'// For Credit Check
'//-----------------------------------------------------------------------------------------------

Case 2
        frmCashier.ListView2.ListItems(pIndex).SubItems(5) = Format(Me.txtCheckAmount, "###,##0.00")
        frmCashier.ListView2.ListItems(pIndex).SubItems(8) = Me.txtCheckNumber
        frmCashier.ListView2.ListItems(pIndex).SubItems(9) = Me.txtCheckBank
        frmCashier.ListView2.ListItems(pIndex).SubItems(10) = Me.txtCheckBranch
        frmCashier.ListView2.ListItems(pIndex).SubItems(14) = Format(Now, "mm/dd/yyyy")
        frmCashier.ListView2.ListItems(pIndex).SubItems(15) = Format(Me.txtCheckDate, "mm/dd/yyyy")
        frmCashier.ListView2.ListItems(pIndex).SubItems(16) = IIf(Me.OptYes, "Yes", IIf(Me.OptNo, "No", "Yes"))
        frmCashier.ListView2.ListItems(pIndex).SubItems(17) = Me.txtNoofDays
        Call GroupBank_Account(21, pIndex)
        
Case 3

'//-----------------------------------------------------------------------------------------------
'// For Credit others
'//-----------------------------------------------------------------------------------------------

        frmCashier.ListView2.ListItems(pIndex).SubItems(7) = Format(Me.txtOthersAmount, "###,##0.00")
Case Else
     MsgBox "Error cannot determine which payment selected", vbInformation
End Select

'//----------------------------
'// For every Statement Balance
'//----------------------------
frmCashier.ListView2.ListItems(pIndex).SubItems(18) = _
                                                    Format(CDbl(frmCashier.ListView2.ListItems(pIndex).SubItems(1)) - _
                                                    (CDbl(frmCashier.ListView2.ListItems(pIndex).SubItems(4)) + _
                                                    CDbl(frmCashier.ListView2.ListItems(pIndex).SubItems(5)) + _
                                                    CDbl(frmCashier.ListView2.ListItems(pIndex).SubItems(6)) + _
                                                    CDbl(frmCashier.ListView2.ListItems(pIndex).SubItems(7))), "###,##0.00")



Call frmCashier.ReCompute
Unload Me
End Sub


Private Sub Form_Activate()
OldBal = frmCashier.ListView2.ListItems(CDbl(Me.Tag)).SubItems(1)
End Sub

Private Sub Form_Load()
TmpSel = 0
myCardFlag = False

End Sub

Private Sub imgXpTab1_Click()
'SD105-102500616-1
Dim mytmpbal
Dim tmpmarkup
TmpSel = Me.imgXpTab1.ActiveTab

frmCashier.ListView2.ListItems(CDbl(Me.Tag)).SubItems(1) = Format(OldBal, "###,##0.00")
If CDbl(TmpSel) = 1 Then

        mytmpbal = frmCashier.ListView2.ListItems(CDbl(Me.Tag)).SubItems(1)
        tmpmarkup = CDbl(mytmpbal) * (CDbl(txtMarkPercent) / 100)

    Me.txtCardAmount = Format(mytmpbal + tmpmarkup, "###,##0.00")
Else
    Me.txtCardAmount = "0.00"
End If
If CDbl(TmpSel) = 2 Then
    Me.txtCheckDate = Format(Now, "mm/dd/yyyy")
End If
End Sub


Private Sub txtCardAmount_GotFocus()
pIndex = CDbl(Me.Tag)
If CDbl(TmpSel) = 1 Then
        tmpBal = frmCashier.ListView2.ListItems(pIndex).SubItems(1)
        tmpmarkup = CDbl(tmpBal) * (CDbl(txtMarkPercent) / 100)
        frmCashier.ListView2.ListItems(pIndex).SubItems(1) = Format(CDbl(tmpBal) + tmpmarkup, "###,##0.00")
End If
End Sub

Private Sub txtMarkPercent_Change()
Dim mytmpbal
Dim mytmpmarkup
If IsNumeric(Me.txtMarkPercent) Then
   mytmpbal = frmCashier.ListView2.ListItems(CDbl(Me.Tag)).SubItems(1)
   tmpmarkup = CDbl(mytmpbal) * (CDbl(txtMarkPercent) / 100)
   Me.txtCardAmount = Format(mytmpbal + tmpmarkup, "###,##0.00")
Else
    Me.txtMarkPercent = "0.00"
    Call kulotHL(Me.txtMarkPercent)
End If
End Sub
