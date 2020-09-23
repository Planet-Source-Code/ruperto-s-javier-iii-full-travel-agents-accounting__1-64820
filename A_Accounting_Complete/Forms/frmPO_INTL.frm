VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPO_INTL 
   Caption         =   "Purchase Order International"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Select Payment"
      Height          =   1185
      Left            =   3795
      TabIndex        =   89
      Top             =   5790
      Width           =   3075
      Begin VB.TextBox txtExchangeRate 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1500
         TabIndex        =   92
         Text            =   "0.00"
         Top             =   705
         Width           =   1395
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Dollar"
         Height          =   330
         Left            =   1545
         TabIndex        =   91
         Top             =   330
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Peso"
         Height          =   270
         Left            =   300
         TabIndex        =   90
         Top             =   345
         Width           =   1185
      End
      Begin VB.Label Label32 
         Caption         =   "Exchange Rate :"
         Height          =   285
         Left            =   225
         TabIndex        =   93
         Top             =   735
         Width           =   1350
      End
   End
   Begin VB.TextBox txtIssuedBy 
      Enabled         =   0   'False
      Height          =   345
      Left            =   10065
      TabIndex        =   87
      Top             =   7860
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPO_INTL.frx":0000
      Left            =   1035
      List            =   "frmPO_INTL.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   83
      Top             =   5880
      Width           =   2745
   End
   Begin VB.Frame Frame2 
      Caption         =   "Command"
      Height          =   765
      Left            =   150
      TabIndex        =   77
      Top             =   8265
      Width           =   12810
      Begin VB.CommandButton cmdOVerRide 
         Caption         =   "Over-Ride"
         Height          =   390
         Left            =   5355
         TabIndex        =   85
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Delete"
         Height          =   390
         Left            =   4050
         TabIndex        =   86
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Recalc"
         Height          =   390
         Left            =   7695
         TabIndex        =   82
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   390
         Left            =   11340
         TabIndex        =   81
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   390
         Left            =   2745
         TabIndex        =   80
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Enabled         =   0   'False
         Height          =   390
         Left            =   1440
         TabIndex        =   79
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   390
         Left            =   135
         TabIndex        =   78
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PURCHASE ORDER"
      Height          =   1740
      Left            =   60
      TabIndex        =   73
      Top             =   120
      Width           =   6840
      Begin VB.TextBox txtSAno 
         Height          =   345
         Left            =   2310
         TabIndex        =   2
         Top             =   1170
         Width           =   2685
      End
      Begin VB.TextBox txtPONumber 
         Height          =   345
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   765
         Width           =   2685
      End
      Begin VB.TextBox txtDate 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2310
         TabIndex        =   0
         Top             =   375
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "STATEMENT # :"
         Height          =   315
         Left            =   210
         TabIndex        =   76
         Top             =   1230
         Width           =   2070
      End
      Begin VB.Label Label2 
         Caption         =   "PO # :"
         Height          =   315
         Left            =   225
         TabIndex        =   75
         Top             =   765
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "DATE :"
         Height          =   315
         Left            =   225
         TabIndex        =   74
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.TextBox txtParticular 
      Height          =   345
      Left            =   1485
      TabIndex        =   9
      Top             =   7395
      Width           =   5160
   End
   Begin VB.Frame FrameBasic 
      Caption         =   "Basic"
      Height          =   4110
      Left            =   6975
      TabIndex        =   49
      Top             =   0
      Width           =   5985
      Begin VB.TextBox txtMiscName 
         Height          =   345
         Left            =   90
         TabIndex        =   24
         Top             =   2910
         Width           =   2430
      End
      Begin VB.TextBox txtSubTot_Basic 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Index           =   1
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0.00"
         Top             =   3630
         Width           =   1305
      End
      Begin VB.TextBox txtSubTot_Basic 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Index           =   0
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "0.00"
         Top             =   3630
         Width           =   1305
      End
      Begin VB.TextBox txtPhilTravelTax 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   3270
         Width           =   1305
      End
      Begin VB.TextBox txtPhilTravelTax 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   3270
         Width           =   1305
      End
      Begin VB.TextBox txtMisc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   2910
         Width           =   1305
      End
      Begin VB.TextBox txtMisc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   2910
         Width           =   1305
      End
      Begin VB.TextBox txtDomestic 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2550
         Width           =   1305
      End
      Begin VB.TextBox txtDomestic 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   2550
         Width           =   1305
      End
      Begin VB.TextBox txttaxIns 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2190
         Width           =   1305
      End
      Begin VB.TextBox txttaxIns 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   2190
         Width           =   1305
      End
      Begin VB.TextBox txtcar 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   1830
         Width           =   1305
      End
      Begin VB.TextBox txtcar 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   1830
         Width           =   1305
      End
      Begin VB.TextBox txtHotel 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   1470
         Width           =   1305
      End
      Begin VB.TextBox txtHotel 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   1470
         Width           =   1305
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1110
         Width           =   1305
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1110
         Width           =   1305
      End
      Begin VB.TextBox txtVUSA 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox txtVUSA 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   375
         Width           =   1305
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   390
         Width           =   1305
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "PESO"
         Height          =   315
         Left            =   4530
         TabIndex        =   63
         Top             =   135
         Width           =   1290
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "DOLLAR"
         Height          =   315
         Left            =   3075
         TabIndex        =   62
         Top             =   135
         Width           =   1275
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "SUB-TOTAL :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   59
         Top             =   3615
         Width           =   2760
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "PHIL. TRAVEL TAX :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   58
         Top             =   3255
         Width           =   2760
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "MISC :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   57
         Top             =   2895
         Width           =   2760
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "DOMESTIC TICKET :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   56
         Top             =   2535
         Width           =   2760
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "TAX INSURANCE :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   55
         Top             =   2175
         Width           =   2760
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "CAR RENTAL :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   54
         Top             =   1830
         Width           =   2760
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "HOTEL ACCOMODATION :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   53
         Top             =   1470
         Width           =   2760
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "SWING AROUND / PALAKBAYAN :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   52
         Top             =   1110
         Width           =   2760
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "VUSA :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   51
         Top             =   735
         Width           =   2760
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "FARE :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   50
         Top             =   375
         Width           =   2760
      End
   End
   Begin VB.Frame FrameLess 
      Caption         =   "Deductions / Commissions"
      Height          =   3675
      Left            =   6990
      TabIndex        =   44
      Top             =   4110
      Width           =   5985
      Begin VB.TextBox txtGrandTot_all 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   0
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "0.00"
         Top             =   3150
         Width           =   1305
      End
      Begin VB.TextBox txtGrandTot_all 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   1
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   3150
         Width           =   1305
      End
      Begin VB.TextBox txtGrandTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   1
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   2775
         Width           =   1305
      End
      Begin VB.TextBox txtGrandTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   0
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0.00"
         Top             =   2775
         Width           =   1305
      End
      Begin VB.TextBox txtSubTot_Deduc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   1290
         Width           =   1305
      End
      Begin VB.TextBox txtSubTot_Deduc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   1290
         Width           =   1305
      End
      Begin VB.TextBox txtEvat 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   2235
         Width           =   1305
      End
      Begin VB.TextBox txtEvat 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   2235
         Width           =   1305
      End
      Begin VB.TextBox txtSwing_Less 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   4530
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox txtSwing_Less 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox txtCommEvat 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   32
         Text            =   "0"
         Top             =   2235
         Width           =   1635
      End
      Begin VB.TextBox txtCommSwing 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   285
         TabIndex        =   29
         Text            =   "0"
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label lblPassCtr 
         Caption         =   "1"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1950
         TabIndex        =   95
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label33 
         Caption         =   "GRAND TOTAL :    x  "
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   300
         TabIndex        =   94
         Top             =   3240
         Width           =   1680
      End
      Begin VB.Label Label28 
         Caption         =   "%"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1950
         TabIndex        =   67
         Top             =   2235
         Width           =   690
      End
      Begin VB.Label Label27 
         Caption         =   "%"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1920
         TabIndex        =   66
         Top             =   975
         Width           =   690
      End
      Begin VB.Label Label24 
         Caption         =   "TOTAL :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   61
         Top             =   2865
         Width           =   915
      End
      Begin VB.Label Label23 
         Caption         =   "SUB-TOTAL :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   60
         Top             =   1395
         Width           =   1980
      End
      Begin VB.Label Label22 
         Caption         =   "EVAT"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   48
         Top             =   1965
         Width           =   1680
      End
      Begin VB.Label Label21 
         Caption         =   "----------ADD ---------"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   285
         TabIndex        =   47
         Top             =   1740
         Width           =   2610
      End
      Begin VB.Label Label20 
         Caption         =   "COMM ON SWING AROUND / PALAKBAYAN"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   46
         Top             =   660
         Width           =   3570
      End
      Begin VB.Label Label19 
         Caption         =   "---------- LESS ---------"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   255
         TabIndex        =   45
         Top             =   330
         Width           =   2610
      End
   End
   Begin VB.TextBox txtFOC 
      Height          =   345
      Left            =   1485
      TabIndex        =   8
      Top             =   7005
      Width           =   3210
   End
   Begin VB.CommandButton cmdRemovePax 
      Caption         =   "Remove PAX"
      Height          =   465
      Left            =   1965
      TabIndex        =   42
      Top             =   6345
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddPax 
      Caption         =   "Add PAX"
      Height          =   465
      Left            =   135
      TabIndex        =   41
      Top             =   6345
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   3750
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POIDDetails"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "POID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Passenger Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ticket No"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtOthers 
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   3165
      Width           =   4950
   End
   Begin VB.TextBox txtRecordLoc 
      Height          =   345
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   4950
   End
   Begin VB.TextBox txtRoute 
      Height          =   345
      Left            =   810
      TabIndex        =   4
      Top             =   2295
      Width           =   6060
   End
   Begin VB.TextBox txtTo 
      Height          =   345
      Left            =   810
      TabIndex        =   3
      Top             =   1890
      Width           =   6060
   End
   Begin VB.Label Label31 
      Caption         =   "ISSUED BY :"
      Height          =   315
      Left            =   8985
      TabIndex        =   88
      Top             =   7875
      Width           =   1065
   End
   Begin VB.Label Label30 
      Caption         =   "STATUS :"
      Height          =   285
      Left            =   150
      TabIndex        =   84
      Top             =   5895
      Width           =   945
   End
   Begin VB.Label Label29 
      Caption         =   "PARTICULARS :"
      Height          =   315
      Left            =   195
      TabIndex        =   72
      Top             =   7425
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "FOC # :"
      Height          =   315
      Left            =   195
      TabIndex        =   43
      Top             =   7020
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "OTHERS :"
      Height          =   315
      Left            =   135
      TabIndex        =   40
      Top             =   3165
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "RECORD LOCATOR :"
      Height          =   315
      Left            =   150
      TabIndex        =   39
      Top             =   2790
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "ROUTE"
      Height          =   315
      Left            =   165
      TabIndex        =   38
      Top             =   2295
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "TO :"
      Height          =   315
      Left            =   150
      TabIndex        =   37
      Top             =   1905
      Width           =   600
   End
End
Attribute VB_Name = "frmPO_INTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngIndex    As Long
Dim SQL         As String
Private Sub cmdAddPax_Click()
frmPO_INTL_insert.Show 1
End Sub

Private Sub cmdCancel_Click()
Me.ListView1.ListItems.Clear
Call kulotClrText(Me)
txtSubTot_Basic(0) = "0.00"
txtSubTot_Basic(1) = "0.00"
txtGrandTot(0) = "0.00"
txtGrandTot(1) = "0.00"
Me.txtDate = RetDate(Now)

If ReturnFirst(GetLastNumber) = 0 Then
    Me.txtPONumber = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
Else
    Me.txtPONumber = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
End If

With Me
    .cmdNew.Enabled = True
    .cmdPost.Enabled = False
    .cmdCancel.Enabled = False
    .cmdFind.Enabled = False
    .cmdOVerRide.Enabled = True
    .Frame1 = "PURCHASE ORDER"
End With

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo FailSafe_Err
Dim ask As Integer

If Not FindPO(Me.txtPONumber) Then
    MsgBox "This PO :" & Me.txtPONumber & " was not save cannot delete", vbInformation
    Exit Sub
End If

ask = MsgBox("Sure you want to delete this PO?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub
        If Not CheckNull(Me.txtPONumber) Then
        cn.BeginTrans
                    SQL = "DELETE * FROM Tbl_PO_INTL WHERE [Po Number]='" & Me.txtPONumber & "'"
                    cn.Execute SQL
        cn.CommitTrans
            MsgBox "PO Successfully deleted", vbInformation
            Call cmdCancel_Click
                With Me
                    .cmdNew.Enabled = True
                    .cmdPost.Enabled = False
                    .cmdCancel.Enabled = False
                    .cmdFind.Enabled = True
                    .cmdOVerRide.Enabled = True
                End With
        End If
Exit Sub
FailSafe_Err:
cn.RollbackTrans
MsgBox "There was an error while deleting the PO " & Err.Description
End Sub

Private Sub cmdNew_Click()
Me.ListView1.ListItems.Clear
Call kulotClrText(Me)
txtSubTot_Basic(0) = "0.00"
txtSubTot_Basic(1) = "0.00"
txtGrandTot(0) = "0.00"
txtGrandTot(1) = "0.00"
Me.txtDate = RetDate(Now)
Me.txtIssuedBy = MDImain.StatusBar1.Panels(2).Text

If ReturnFirst(GetLastNumber) = 0 Then
    Me.txtPONumber = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
Else
    Me.txtPONumber = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
End If

Me.txtTo.SetFocus
With Me
    .cmdNew.Enabled = False
    .cmdPost.Enabled = True
    .cmdCancel.Enabled = True
    .cmdFind.Enabled = False
    .cmdOVerRide.Enabled = False
    .Frame1 = "PURCHASE ORDER"
End With

End Sub

Private Sub cmdOverRide_Click()
frmPO_INTL_Find.Show 1
End Sub

Function FindPO(Param) As Boolean
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM Tbl_PO_INTL WHERE [PO Number]='" & Param & "'"

With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
            FindPO = True
            Else
            FindPO = False
            End If
                        .Close
            Set Rst = Nothing

End With
End Function

Private Sub cmdPost_Click()
'On Error GoTo FailSafe_Err
Dim Rst         As New ADODB.Recordset
Dim RstDet      As New ADODB.Recordset
Dim i           As Integer
Dim ask         As Integer


If Me.ListView1.ListItems.Count < 1 Then
    MsgBox "Please add PAX to post this PO", vbInformation
    Exit Sub
End If

ask = MsgBox("Sure to save this PO?", vbYesNo + vbInformation)
If ask = vbNo Then: Exit Sub
Call Recalc

If UCase(Me.Frame1.Caption) = "OVER-RIDE" Then
    If FindPO(Me.txtPONumber) Then
        SQL = "DELETE * FROM Tbl_PO_INTL WHERE [Po Number]='" & Me.txtPONumber & "'"
        cn.BeginTrans
            cn.Execute SQL
        cn.CommitTrans
    End If
End If

SQL = "SELECT * FROM Tbl_PO_INTL"
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
         cn.BeginTrans
           .AddNew
            ![Po Number] = Me.txtPONumber
            ![PO Date] = Me.txtDate
            ![PO SAno] = Me.txtSAno
            ![Pay to] = Me.txtTo
            !Route = Me.txtRoute
            ![Record Locator] = Me.txtRecordLoc
            !Others = Me.txtOthers
            ![FOC no] = Me.txtFOC
            !Particulars = Me.txtParticular
            ![Fare Dollar] = Me.txtFare(0)
            ![Fare Peso] = Me.txtFare(1)
            ![VUSA Dollar] = Me.txtVUSA(0)
            ![VUSA Peso] = Me.txtVUSA(1)
            ![Swing Around Dollar] = Me.txtSwing(0)
            ![Swing Around Peso] = Me.txtSwing(1)
            ![Hotel Acco Dollar] = Me.txtHotel(0)
            ![Hotel Acco Peso] = Me.txtHotel(1)
            ![Car Rental Dollar] = Me.txtcar(0)
            ![Car Rental Peso] = Me.txtcar(1)
            ![Tax Ins Dollar] = Me.txttaxIns(0)
            ![Tax Ins Peso] = Me.txttaxIns(1)
            ![Domestic Ticket Dollar] = Me.txtDomestic(0)
            ![Domestic Ticket Peso] = Me.txtDomestic(1)
            ![Misc Name] = Me.txtMiscName
            ![Misc Amount Dollar] = Me.txtMisc(0)
            ![Misc Amount Peso] = Me.txtMisc(1)
            ![Phil Travel Tax Dollar] = Me.txtPhilTravelTax(0)
            ![Phil Travel Tax Peso] = Me.txtPhilTravelTax(1)
            ![SubTot Basic Dollar] = Me.txtSubTot_Basic(0)
            ![SubTot Basic Peso] = Me.txtSubTot_Basic(1)
            ![Swing Comm Percent] = Me.txtCommSwing
            ![Swing Comm Dollar] = Me.txtSwing_Less(0)
            ![Swing Comm Peso] = Me.txtSwing_Less(1)
            ![Evat Percent] = Me.txtCommEvat
            ![Evat Dollar] = Me.txtEvat(0)
            ![Evat Peso] = Me.txtEvat(1)
            ![Total Peso] = Me.txtGrandTot(1)
            ![Total Dollar] = Me.txtGrandTot(0)
            ![SubTot Deduc Dollar] = Me.txtSubTot_Deduc(0)
            ![SubTot Deduc Peso] = Me.txtSubTot_Deduc(1)
            ![Grand Total Dollar] = Me.txtGrandTot_all(0)
            ![Grand Total Peso] = Me.txtGrandTot_all(1)
            ![POSTED] = Me.Combo1
            ![Issued By] = Me.txtIssuedBy
            ![Exchange Rate] = CDbl(Me.txtExchangeRate)
           .Update
           
    If Me.ListView1.ListItems.Count > 0 Then
           SQL = "SELECT * FROM Tbl_PO_INTL_Details"
           RstDet.Open SQL, cn, adOpenKeyset, adLockOptimistic
        For i = 1 To Me.ListView1.ListItems.Count
           RstDet.AddNew
                RstDet![PoID] = .Fields("POID").Value
                RstDet![Po Number] = .Fields("Po Number").Value
                RstDet![Pax Name] = Me.ListView1.ListItems(i).SubItems(2)
                RstDet![Ticket No] = Me.ListView1.ListItems(i).SubItems(3)
           RstDet.Update
        Next i
    End If
         cn.CommitTrans
        If UCase(Me.Frame1.Caption) = "OVER-RIDE" Then
         MsgBox "PO International successfully updated...", vbInformation
        Else
         MsgBox "PO International successfully save...", vbInformation
        End If
       Rst.Close
     Set Rst = Nothing
     Call cmdNew_Click
End With

With Me
    .cmdNew.Enabled = True
    .cmdPost.Enabled = False
    .cmdCancel.Enabled = False
    .cmdFind.Enabled = True
    .cmdOVerRide.Enabled = True
End With

Exit Sub
FailSafe_Err:
cn.RollbackTrans
MsgBox "An error occured while saving the PO " & Err.Description, vbInformation
End Sub

Private Sub cmdRecalc_Click()
Call Recalc
End Sub

Private Sub cmdRemovePax_Click()
If Me.ListView1.ListItems.Count > 0 Then
    Me.ListView1.ListItems.Remove lngIndex
End If
End Sub

Private Sub Form_Load()
Me.txtDate = RetDate(Now)
Me.txtIssuedBy = MDImain.StatusBar1.Panels(2).Text
End Sub

Function GetLastNumber() As String
Dim RsFnumber       As ADODB.Recordset
Dim SQL             As String
Dim Tmp             As String
Dim myTmpPos        As Integer

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT [Po Number] from Tbl_PO_INTL ORDER by [Po Number] ASC"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               Tmp = RsFnumber("Po Number").Value
               myTmpPos = Int(ReturnFirst(Tmp)) - (Int(ReturnFirst(Tmp)) - Int(Return_1stDash(Tmp)))
               Tmp = Mid(Tmp, Return_1stDash(Tmp) + 5, myTmpPos - 1)
               GetLastNumber = "POI" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & AutoIncrement(Tmp)
        Else
               GetLastNumber = "POI" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & "000000000"
        End If
End With

End Function

Function Recalc() As Double
Dim TmpBasic(1 To 2)    As Double
Dim TmpDeduc(1 To 2)    As Double
Dim TmpEvat(1 To 2)     As Double
Dim TmpGrandTot(1 To 2) As Double
Dim PassCtr             As Integer
'Index 1 are dollars
'Index 2 are peso

       TmpBasic(1) = _
           CDbl(Me.txtFare(0)) + _
           CDbl(Me.txtVUSA(0)) + _
           CDbl(Me.txtSwing(0)) + _
           CDbl(Me.txtHotel(0)) + _
           CDbl(Me.txtcar(0)) + _
           CDbl(Me.txttaxIns(0)) + _
           CDbl(Me.txtDomestic(0)) + _
           CDbl(Me.txtMisc(0)) + _
           CDbl(Me.txtPhilTravelTax(0))

           
        TmpBasic(2) = _
           CDbl(Me.txtFare(1)) + _
           CDbl(Me.txtVUSA(1)) + _
           CDbl(Me.txtSwing(1)) + _
           CDbl(Me.txtHotel(1)) + _
           CDbl(Me.txtcar(1)) + _
           CDbl(Me.txttaxIns(1)) + _
           CDbl(Me.txtDomestic(1)) + _
           CDbl(Me.txtMisc(1)) + _
           CDbl(Me.txtPhilTravelTax(1))
           
        TmpDeduc(1) = (CDbl(txtCommSwing) / 100) * CDbl(Me.txtSwing(0))
        TmpDeduc(2) = (CDbl(txtCommSwing) / 100) * CDbl(Me.txtSwing(1))
        
        TmpEvat(1) = CDbl(Me.txtSubTot_Basic(0)) * (CDbl(txtCommEvat) / 100)
        TmpEvat(2) = CDbl(Me.txtSubTot_Basic(1)) * (CDbl(txtCommEvat) / 100)
        
        Me.txtSubTot_Basic(0) = RetCurrency(TmpBasic(1))
        Me.txtSubTot_Basic(1) = RetCurrency(TmpBasic(2))
        
       
        Me.txtSwing_Less(0) = RetCurrency(TmpDeduc(1))
        Me.txtSwing_Less(1) = RetCurrency(TmpDeduc(2))
       
        txtSubTot_Deduc(0) = RetCurrency((TmpBasic(1) - TmpDeduc(1)))
        txtSubTot_Deduc(1) = RetCurrency((TmpBasic(2) - TmpDeduc(2)))
        
        
        Me.txtEvat(0) = RetCurrency(TmpEvat(1))
        Me.txtEvat(1) = RetCurrency(TmpEvat(2))
        
        PassCtr = Me.ListView1.ListItems.Count
        Me.lblPassCtr = PassCtr
        
        TmpGrandTot(1) = CDbl(txtSubTot_Deduc(0)) + TmpEvat(1)
        TmpGrandTot(2) = CDbl(txtSubTot_Deduc(1)) + TmpEvat(2)
        
        
        
        Me.txtGrandTot(0) = RetCurrency(TmpGrandTot(1))
        Me.txtGrandTot(1) = RetCurrency(TmpGrandTot(2))
        
        
        Me.txtGrandTot_all(0) = RetCurrency(CDbl(TmpGrandTot(1)) * PassCtr)
        Me.txtGrandTot_all(1) = RetCurrency(CDbl(TmpGrandTot(2)) * PassCtr)
   
End Function

Sub LoadValues(Param)
Dim Rst                 As New ADODB.Recordset
Dim RsPODetails         As New ADODB.Recordset
Dim mySQL               As String
Dim ctr                 As Integer

  
SQL = "SELECT * from Tbl_PO_INTL WHERE [PoID]=" & Param & " ORDER by [Po Number] ASC"

With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Me.txtPONumber = ![Po Number]
            Me.txtDate = ![PO Date]
            Me.txtSAno = ![PO SAno]
            Me.txtTo = ![Pay to]
            Me.txtRoute = !Route
            Me.txtRecordLoc = ![Record Locator]
            Me.txtOthers = !Others
            Me.txtFOC = ![FOC no]
            Me.txtParticular = !Particulars
            Me.txtFare(0) = RetCurrency(![Fare Dollar])
            Me.txtFare(1) = RetCurrency(![Fare Peso])
            Me.txtVUSA(0) = RetCurrency(![VUSA Dollar])
            Me.txtVUSA(1) = RetCurrency(![VUSA Peso])
            Me.txtSwing(0) = RetCurrency(![Swing Around Dollar])
            Me.txtSwing(1) = RetCurrency(![Swing Around Peso])
            Me.txtHotel(0) = RetCurrency(![Hotel Acco Dollar])
            Me.txtHotel(1) = RetCurrency(![Hotel Acco Peso])
            Me.txtcar(0) = RetCurrency(![Car Rental Dollar])
            Me.txtcar(1) = RetCurrency(![Car Rental Peso])
            Me.txttaxIns(0) = RetCurrency(![Tax Ins Dollar])
            Me.txttaxIns(1) = RetCurrency(![Tax Ins Peso])
            Me.txtDomestic(0) = RetCurrency(![Domestic Ticket Dollar])
            Me.txtDomestic(1) = RetCurrency(![Domestic Ticket Peso])
            Me.txtMiscName = RetCurrency(![Misc Name])
            Me.txtMisc(0) = RetCurrency(![Misc Amount Dollar])
            Me.txtMisc(1) = RetCurrency(![Misc Amount Peso])
            Me.txtPhilTravelTax(0) = RetCurrency(![Phil Travel Tax Dollar])
            Me.txtPhilTravelTax(1) = RetCurrency(![Phil Travel Tax Peso])
            Me.txtSubTot_Basic(0) = RetCurrency(![SubTot Basic Dollar])
            Me.txtSubTot_Basic(1) = RetCurrency(![SubTot Basic Peso])
            Me.txtCommSwing = RetCurrency(![Swing Comm Percent])
            Me.txtSwing_Less(0) = RetCurrency(![Swing Comm Dollar])
            Me.txtSwing_Less(1) = RetCurrency(![Swing Comm Peso])
            Me.txtCommEvat = RetCurrency(![Evat Percent])
            Me.txtEvat(0) = RetCurrency(![Evat Dollar])
            Me.txtEvat(1) = RetCurrency(![Evat Peso])
            Me.txtSubTot_Deduc(0) = RetCurrency(![SubTot Deduc Dollar])
            Me.txtSubTot_Deduc(1) = RetCurrency(![SubTot Deduc Peso])
            
            Me.txtGrandTot(0) = RetCurrency(![Total Dollar])
            Me.txtGrandTot(1) = RetCurrency(![Total Peso])
            Me.txtGrandTot_all(0) = RetCurrency(![Grand Total Dollar])
            Me.txtGrandTot_all(1) = RetCurrency(![Grand Total Peso])
            
            Me.txtIssuedBy = RetCurrency(![Issued By])
       If .Fields("Posted").Value = "POSTED" Then
                Me.Combo1.ListIndex = 0
            Else
                Me.Combo1.ListIndex = 1
       End If
       Me.txtExchangeRate = IIf(Not IsNull(![Exchange Rate]), ![Exchange Rate], "0.00")
         Me.Tag = .Fields("PoID").Value

'//=======================================================================
'//pull out data from details and load it to list view
'//=======================================================================
mySQL = "SELECT * FROM Tbl_PO_INTL_Details WHERE [POid]=" & Param
Me.ListView1.ListItems.Clear
         With RsPODetails
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then

                .MoveFirst
                    ctr = 0
                    On Error Resume Next
                    Do While Not .EOF
                    ctr = ctr + 1
                        ListView1.ListItems.Add , , .Fields("POIDDetails").Value
                        ListView1.ListItems.Item(ctr).SubItems(1) = .Fields("POID").Value
                        ListView1.ListItems.Item(ctr).SubItems(2) = .Fields("Pax Name").Value
                        ListView1.ListItems.Item(ctr).SubItems(3) = .Fields("Ticket No").Value
                        .MoveNext
                    Loop
               End If
        End With


        End If
End With

End Sub

Private Sub ListView1_Click()
If Me.ListView1.ListItems.Count > 0 Then
    lngIndex = Me.ListView1.SelectedItem.Index
End If
End Sub



Private Sub txtcar_Change(Index As Integer)
If Not IsNumeric(Me.txtcar(Index)) Then
    Me.txtcar(Index) = "0.00"
    Call kulotHL(Me.txtcar(Index))
End If
If Me.Option1 Then
    Me.txtcar(1) = Me.txtcar(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txtCommEvat_Change()
If Not IsNumeric(Me.txtCommEvat) Then
    Me.txtCommEvat = "0.00"
    Call kulotHL(Me.txtCommEvat)
End If
End Sub

Private Sub txtCommSwing_Change()
If Not IsNumeric(Me.txtCommSwing) Then
    Me.txtCommSwing = "0.00"
    Call kulotHL(Me.txtCommSwing)
End If
End Sub

Private Sub txtDomestic_Change(Index As Integer)
If Not IsNumeric(Me.txtDomestic(Index)) Then
    Me.txtDomestic(Index) = "0.00"
    Call kulotHL(Me.txtDomestic(Index))
End If
If Me.Option1 Then
    Me.txtDomestic(1) = Me.txtDomestic(0) * CDbl(Me.txtExchangeRate)
End If

End Sub

Private Sub txtEvat_Change(Index As Integer)
If Not IsNumeric(Me.txtEvat(Index)) Then
    Me.txtEvat(Index) = "0.00"
    Call kulotHL(Me.txtEvat(Index))
End If
If Me.Option1 Then
    Me.txtEvat(1) = Me.txtEvat(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txtExchangeRate_Change()
If Not IsNumeric(Me.txtExchangeRate) Then
    Me.txtExchangeRate = "0.00"
    Call kulotHL(Me.txtExchangeRate)
End If
End Sub

Private Sub txtFare_Change(Index As Integer)
If Not IsNumeric(Me.txtFare(Index)) Then
    Me.txtFare(Index) = "0.00"
    Call kulotHL(Me.txtFare(Index))
End If
If Me.Option1 Then
    Me.txtFare(1) = Me.txtFare(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txtHotel_Change(Index As Integer)
If Not IsNumeric(Me.txtHotel(Index)) Then
    Me.txtHotel(Index) = "0.00"
    Call kulotHL(Me.txtHotel(Index))
End If
If Me.Option1 Then
    Me.txtHotel(1) = Me.txtHotel(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txtMisc_Change(Index As Integer)
If Not IsNumeric(Me.txtMisc(Index)) Then
    Me.txtMisc(Index) = "0.00"
    Call kulotHL(Me.txtMisc(Index))
End If
If Me.Option1 Then
    Me.txtMisc(1) = Me.txtMisc(0) * CDbl(Me.txtExchangeRate)
End If

End Sub

Private Sub txtPhilTravelTax_Change(Index As Integer)
If Not IsNumeric(Me.txtPhilTravelTax(Index)) Then
    Me.txtPhilTravelTax(Index) = "0.00"
    Call kulotHL(Me.txtPhilTravelTax(Index))
End If
If Me.Option1 Then
    Me.txtPhilTravelTax(1) = Me.txtPhilTravelTax(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txtSubTot_Deduc_Change(Index As Integer)
If Not IsNumeric(Me.txtSubTot_Deduc(Index)) Then
    Me.txtSubTot_Deduc(Index) = "0.00"
    Call kulotHL(Me.txtSubTot_Deduc(Index))
End If
End Sub

Private Sub txtSwing_Change(Index As Integer)
If Not IsNumeric(Me.txtSwing(Index)) Then
    Me.txtSwing(Index) = "0.00"
    Call kulotHL(Me.txtSwing(Index))
End If
If Me.Option1 Then
    Me.txtSwing(1) = Me.txtSwing(0) * CDbl(Me.txtExchangeRate)
End If


End Sub

Private Sub txtSwing_Less_Change(Index As Integer)
If Not IsNumeric(Me.txtSwing_Less(Index)) Then
    Me.txtSwing_Less(Index) = "0.00"
    Call kulotHL(Me.txtSwing_Less(Index))
End If
If Me.Option1 Then
    Me.txtSwing_Less(1) = Me.txtSwing_Less(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txttaxIns_Change(Index As Integer)
If Not IsNumeric(Me.txttaxIns(Index)) Then
    Me.txttaxIns(Index) = "0.00"
    Call kulotHL(Me.txttaxIns(Index))
End If
If Me.Option1 Then
    Me.txttaxIns(1) = Me.txttaxIns(0) * CDbl(Me.txtExchangeRate)
End If
End Sub

Private Sub txtVUSA_Change(Index As Integer)
If Not IsNumeric(Me.txtVUSA(Index)) Then
    Me.txtVUSA(Index) = "0.00"
    Call kulotHL(Me.txtVUSA(Index))
End If
If Me.Option1 Then
    Me.txtVUSA(1) = Me.txtVUSA(0) * CDbl(Me.txtExchangeRate)
End If

End Sub
