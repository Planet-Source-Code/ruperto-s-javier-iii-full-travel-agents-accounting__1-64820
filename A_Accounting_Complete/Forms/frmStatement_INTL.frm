VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatement_INTL 
   Caption         =   "Statement International"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditPax 
      Caption         =   "Edit PAX"
      Height          =   465
      Left            =   1605
      TabIndex        =   115
      Top             =   6345
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Payment"
      Height          =   1185
      Left            =   3810
      TabIndex        =   92
      Top             =   5805
      Width           =   3075
      Begin VB.OptionButton Option1 
         Caption         =   "Peso"
         Height          =   270
         Left            =   300
         TabIndex        =   95
         Top             =   345
         Width           =   1185
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Dollar"
         Height          =   330
         Left            =   1545
         TabIndex        =   94
         Top             =   330
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.TextBox txtExchangeRate 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1500
         TabIndex        =   93
         Text            =   "0.00"
         Top             =   705
         Width           =   1395
      End
      Begin VB.Label Label32 
         Caption         =   "Exchange Rate :"
         Height          =   285
         Left            =   225
         TabIndex        =   96
         Top             =   735
         Width           =   1350
      End
   End
   Begin VB.TextBox txtIssuedBy 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1470
      TabIndex        =   85
      Top             =   7890
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmStatement_INTL.frx":0000
      Left            =   1035
      List            =   "frmStatement_INTL.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   81
      Top             =   5880
      Width           =   2745
   End
   Begin VB.Frame Frame2 
      Caption         =   "Command"
      Height          =   765
      Left            =   150
      TabIndex        =   75
      Top             =   8400
      Width           =   12810
      Begin VB.CommandButton cmdOVerRide 
         Caption         =   "Over-Ride"
         Height          =   390
         Left            =   5355
         TabIndex        =   83
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Delete"
         Height          =   390
         Left            =   4050
         TabIndex        =   84
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Recalc"
         Height          =   390
         Left            =   7695
         TabIndex        =   80
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   390
         Left            =   11340
         TabIndex        =   79
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   390
         Left            =   2745
         TabIndex        =   78
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Enabled         =   0   'False
         Height          =   390
         Left            =   1440
         TabIndex        =   77
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   390
         Left            =   135
         TabIndex        =   76
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2130
      Left            =   60
      TabIndex        =   71
      Top             =   120
      Width           =   6840
      Begin VB.TextBox txtAgencyName 
         Height          =   315
         Left            =   2310
         TabIndex        =   90
         Top             =   1695
         Width           =   3870
      End
      Begin VB.ComboBox CboAccountName 
         Height          =   315
         Left            =   2310
         TabIndex        =   88
         Top             =   1350
         Width           =   3885
      End
      Begin VB.CommandButton cmdSearchPO 
         Caption         =   "Search"
         Height          =   360
         Left            =   5010
         TabIndex        =   87
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txtSAno 
         Height          =   345
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   2685
      End
      Begin VB.TextBox txtPONumber 
         Height          =   345
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   2685
      End
      Begin VB.TextBox txtDate 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2310
         TabIndex        =   0
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Agency / Name"
         Height          =   390
         Index           =   2
         Left            =   210
         TabIndex        =   91
         Top             =   1710
         Width           =   1980
      End
      Begin VB.Label Label2 
         Caption         =   "Acc. Name:"
         Height          =   390
         Index           =   0
         Left            =   210
         TabIndex        =   89
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "STATEMENT # :"
         Height          =   315
         Left            =   210
         TabIndex        =   74
         Top             =   1005
         Width           =   2070
      End
      Begin VB.Label Label2 
         Caption         =   "PO # :"
         Height          =   315
         Index           =   1
         Left            =   225
         TabIndex        =   73
         Top             =   540
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "DATE :"
         Height          =   315
         Left            =   225
         TabIndex        =   72
         Top             =   135
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
      Height          =   4290
      Left            =   6975
      TabIndex        =   49
      Top             =   120
      Width           =   8220
      Begin VB.TextBox txtFarePO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   114
         Text            =   "0.00"
         Top             =   585
         Width           =   1215
      End
      Begin VB.TextBox txtVUSAPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   113
         Text            =   "0.00"
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox txtSwingPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   112
         Text            =   "0.00"
         Top             =   1305
         Width           =   1215
      End
      Begin VB.TextBox txtHotelPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   111
         Text            =   "0.00"
         Top             =   1665
         Width           =   1215
      End
      Begin VB.TextBox txtcarPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   110
         Text            =   "0.00"
         Top             =   2025
         Width           =   1215
      End
      Begin VB.TextBox txttaxInsPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   109
         Text            =   "0.00"
         Top             =   2385
         Width           =   1215
      End
      Begin VB.TextBox txtDomesticPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   108
         Text            =   "0.00"
         Top             =   2745
         Width           =   1215
      End
      Begin VB.TextBox txtMiscPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   107
         Text            =   "0.00"
         Top             =   3105
         Width           =   1215
      End
      Begin VB.TextBox txtPhilTravelTaxPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6960
         TabIndex        =   106
         Text            =   "0.00"
         Top             =   3465
         Width           =   1215
      End
      Begin VB.TextBox txtFarePO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   105
         Text            =   "0.00"
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox txtVUSAPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   104
         Text            =   "0.00"
         Top             =   945
         Width           =   1230
      End
      Begin VB.TextBox txtSwingPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   103
         Text            =   "0.00"
         Top             =   1305
         Width           =   1230
      End
      Begin VB.TextBox txtHotelPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   102
         Text            =   "0.00"
         Top             =   1665
         Width           =   1230
      End
      Begin VB.TextBox txtcarPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   101
         Text            =   "0.00"
         Top             =   2025
         Width           =   1230
      End
      Begin VB.TextBox txttaxInsPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   100
         Text            =   "0.00"
         Top             =   2385
         Width           =   1230
      End
      Begin VB.TextBox txtDomesticPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   99
         Text            =   "0.00"
         Top             =   2745
         Width           =   1230
      End
      Begin VB.TextBox txtMiscPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   98
         Text            =   "0.00"
         Top             =   3105
         Width           =   1230
      End
      Begin VB.TextBox txtPhilTravelTaxPO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   4395
         TabIndex        =   97
         Text            =   "0.00"
         Top             =   3465
         Width           =   1230
      End
      Begin VB.TextBox txtMiscName 
         Height          =   345
         Left            =   90
         TabIndex        =   24
         Top             =   3105
         Width           =   2430
      End
      Begin VB.TextBox txtSubTot_Basic 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Index           =   1
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0.00"
         Top             =   3825
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
         Top             =   3825
         Width           =   1305
      End
      Begin VB.TextBox txtPhilTravelTax 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   3465
         Width           =   1305
      End
      Begin VB.TextBox txtPhilTravelTax 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   3465
         Width           =   1305
      End
      Begin VB.TextBox txtMisc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   3105
         Width           =   1305
      End
      Begin VB.TextBox txtMisc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   3105
         Width           =   1305
      End
      Begin VB.TextBox txtDomestic 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox txtDomestic 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   2745
         Width           =   1305
      End
      Begin VB.TextBox txttaxIns 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2385
         Width           =   1305
      End
      Begin VB.TextBox txttaxIns 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   2385
         Width           =   1305
      End
      Begin VB.TextBox txtcar 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   2025
         Width           =   1305
      End
      Begin VB.TextBox txtcar 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   2025
         Width           =   1305
      End
      Begin VB.TextBox txtHotel 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   1665
         Width           =   1305
      End
      Begin VB.TextBox txtHotel 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   1665
         Width           =   1305
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1305
         Width           =   1305
      End
      Begin VB.TextBox txtSwing 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1305
         Width           =   1305
      End
      Begin VB.TextBox txtVUSA 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   945
         Width           =   1305
      End
      Begin VB.TextBox txtVUSA 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   945
         Width           =   1305
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtFare 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   0
         Left            =   3075
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   585
         Width           =   1305
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "PESO"
         Height          =   315
         Left            =   5640
         TabIndex        =   63
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "DOLLAR"
         Height          =   315
         Left            =   3075
         TabIndex        =   62
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "SUB-TOTAL :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   59
         Top             =   3810
         Width           =   2760
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "PHIL. TRAVEL TAX :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   58
         Top             =   3450
         Width           =   2760
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "MISC :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   57
         Top             =   3090
         Width           =   2760
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "DOMESTIC TICKET :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   56
         Top             =   2730
         Width           =   2760
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "TAX INSURANCE :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   55
         Top             =   2370
         Width           =   2760
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "CAR RENTAL :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   54
         Top             =   2025
         Width           =   2760
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "HOTEL ACCOMODATION :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   53
         Top             =   1665
         Width           =   2760
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "SWING AROUND / PALAKBAYAN :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   52
         Top             =   1305
         Width           =   2760
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "VUSA :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   51
         Top             =   930
         Width           =   2760
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "FARE :"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   50
         Top             =   570
         Width           =   2760
      End
   End
   Begin VB.Frame FrameLess 
      Caption         =   "Deductions / Commissions"
      Height          =   3870
      Left            =   6975
      TabIndex        =   44
      Top             =   4410
      Width           =   8220
      Begin VB.TextBox txtGrandTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   0
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "0.00"
         Top             =   2910
         Width           =   1305
      End
      Begin VB.TextBox txtGrandTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   1
         Left            =   5655
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "0.00"
         Top             =   2910
         Width           =   1305
      End
      Begin VB.TextBox txtGrandTot_All 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   1
         Left            =   5655
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   3345
         Width           =   1305
      End
      Begin VB.TextBox txtGrandTot_All 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   0
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0.00"
         Top             =   3345
         Width           =   1305
      End
      Begin VB.TextBox txtSubTot_Deduc 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   1
         Left            =   5655
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
         Left            =   5655
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
         Left            =   5655
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
         Left            =   2100
         TabIndex        =   119
         Top             =   3450
         Width           =   675
      End
      Begin VB.Label Label33 
         Caption         =   "TOTAL (Per Pax):"
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   300
         TabIndex        =   118
         Top             =   2940
         Width           =   1980
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
         Caption         =   "GRAND-TOTAL :      x "
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   285
         TabIndex        =   61
         Top             =   3435
         Width           =   1980
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
      Left            =   2685
      TabIndex        =   42
      Top             =   6345
      Width           =   1065
   End
   Begin VB.CommandButton cmdAddPax 
      Caption         =   "Add PAX"
      Height          =   465
      Left            =   525
      TabIndex        =   41
      Top             =   6345
      Width           =   1065
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1785
      Left            =   120
      TabIndex        =   7
      Top             =   4005
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   3149
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
         Text            =   "SAIDDetails"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SAID"
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
      Top             =   3600
      Width           =   4950
   End
   Begin VB.TextBox txtRecordLoc 
      Height          =   345
      Left            =   1920
      TabIndex        =   5
      Top             =   3195
      Width           =   4950
   End
   Begin VB.TextBox txtRoute 
      Height          =   345
      Left            =   810
      TabIndex        =   4
      Top             =   2730
      Width           =   6060
   End
   Begin VB.TextBox txtTo 
      Height          =   345
      Left            =   810
      TabIndex        =   3
      Top             =   2325
      Width           =   6060
   End
   Begin VB.Label Label31 
      Caption         =   "ISSUED BY :"
      Height          =   315
      Left            =   390
      TabIndex        =   86
      Top             =   7905
      Width           =   1065
   End
   Begin VB.Label Label30 
      Caption         =   "STATUS :"
      Height          =   285
      Left            =   150
      TabIndex        =   82
      Top             =   5895
      Width           =   945
   End
   Begin VB.Label Label29 
      Caption         =   "PARTICULARS :"
      Height          =   315
      Left            =   195
      TabIndex        =   70
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
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "RECORD LOCATOR :"
      Height          =   315
      Left            =   150
      TabIndex        =   39
      Top             =   3225
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "ROUTE"
      Height          =   315
      Left            =   165
      TabIndex        =   38
      Top             =   2730
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "TO :"
      Height          =   315
      Left            =   150
      TabIndex        =   37
      Top             =   2340
      Width           =   600
   End
End
Attribute VB_Name = "frmStatement_INTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngIndex    As Long
Dim SQL         As String

Private Sub CboAccountName_Click()
     txtAgencyName = FindAccountName(Me.CboAccountName)
End Sub
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
Private Sub cmdAddPax_Click()
frmPO_INTL_insert.Tag = "SA_INTL"
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

Me.txtSAno = ""
Me.CboAccountName = Empty

With Me
    .cmdNew.Enabled = True
    .cmdPost.Enabled = False
    .cmdCancel.Enabled = False
    .cmdFind.Enabled = False
    .cmdOVerRide.Enabled = True
   
End With
Me.Frame1.Caption = ""
Me.txtPONumber.SetFocus
End Sub

Private Sub cmdEditPax_Click()
If lngIndex > 0 Then
        With frmPO_INTL_insert
            .Tag = "SA_INTL_EDIT"
            .txtlngIndex = lngIndex
            .txtPaxName = Me.ListView1.ListItems(CLng(lngIndex)).SubItems(2)
            .txtTicketNo = Me.ListView1.ListItems(CLng(lngIndex)).SubItems(3)
            .cmdInsert.Caption = "Edit"
            .Show 1
        End With
Else
    MsgBox "Please Select PAX name by clicking its name", vbInformation
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo FailSafe_Err
Dim ask As Integer

If Not FindSA(Me.txtSAno) Then
    MsgBox "This SA :" & Me.txtSAno & " was not save cannot delete", vbInformation
    Exit Sub
End If

ask = MsgBox("Sure you want to delete this SA?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub
        If Not CheckNull(Me.txtSAno) Then
        cn.BeginTrans
                    SQL = "DELETE * FROM tbl_Statement_INTL WHERE [SAno]='" & Me.txtSAno & "'"
                    cn.Execute SQL
        cn.CommitTrans
            MsgBox "SA Successfully deleted", vbInformation
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
MsgBox "There was an error while deleting the SA " & Err.Description
End Sub

Private Sub cmdNew_Click()
Me.ListView1.ListItems.Clear
Me.Frame1.Caption = ""
Call kulotClrText(Me)
txtSubTot_Basic(0) = "0.00"
txtSubTot_Basic(1) = "0.00"
txtGrandTot(0) = "0.00"
txtGrandTot(1) = "0.00"
Me.txtDate = RetDate(Now)
Me.txtIssuedBy = MDImain.StatusBar1.Panels(2).Text
If ReturnFirst(GetLastNumber) = 0 Then
    Me.txtSAno = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
Else
    Me.txtSAno = GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
End If

Me.txtTo.SetFocus
With Me
    .cmdNew.Enabled = False
    .cmdPost.Enabled = True
    .cmdCancel.Enabled = True
    .cmdFind.Enabled = False
    .cmdOVerRide.Enabled = False
   
End With

End Sub

Private Sub cmdOverRide_Click()
frmStatement_INTL_Find.Show 1
End Sub

Function FindSA(Param) As Boolean
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Statement_INTL WHERE [SAno]='" & Param & "'"

With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
            FindSA = True
            Else
            FindSA = False
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

If CheckNull(Me.CboAccountName) Then
    MsgBox "Please supply account name", vbInformation
    Call kulotHL(Me.CboAccountName)
    Exit Sub
End If

If Me.ListView1.ListItems.Count < 1 Then
    MsgBox "Please add PAX to post this SA", vbInformation
    Exit Sub
End If

ask = MsgBox("Sure to save this Statement?", vbYesNo + vbInformation)
If ask = vbNo Then: Exit Sub

If OK_2_Proceed Then
    Call Recalc
Else
    Exit Sub
End If

If UCase(Me.Frame1.Caption) = "OVER-RIDE" Then
    If FindSA(Me.txtSAno) Then
        SQL = "DELETE * FROM tbl_Statement_INTL WHERE [SAno]='" & Me.txtSAno & "'"
        cn.BeginTrans
            cn.Execute SQL
        cn.CommitTrans
    End If
End If

SQL = "SELECT * FROM tbl_Statement_INTL"
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
         cn.BeginTrans
           .AddNew
            ![Po Number] = Me.txtPONumber
            ![SA Date] = Me.txtDate
            ![SAno] = Me.txtSAno
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
            ![SubTot Deduc Dollar] = Me.txtSubTot_Deduc(0)
            ![SubTot Deduc Peso] = Me.txtSubTot_Deduc(1)
            ![Total Dollar] = Me.txtGrandTot(0)
            ![Total Peso] = Me.txtGrandTot(1)
            ![Grand Total Dollar] = Me.txtGrandTot_All(0)
            ![Grand Total Peso] = Me.txtGrandTot_All(1)
            ![POSTED] = Me.Combo1
            ![Issued By] = Me.txtIssuedBy
            ![Balance] = Me.txtGrandTot(1)
            ![AccountID] = FindAccountID(Me.CboAccountName)
            ![AccountNo] = Me.CboAccountName
            ![AgencyName] = Me.txtAgencyName
            ![Exchange Rate] = Me.txtExchangeRate
           .Update
    
    If Me.ListView1.ListItems.Count > 0 Then
           'Now remove first details to accomodate new entry
           Call RemoveDetails(.Fields("SAID").Value)
           
           SQL = "SELECT * FROM tbl_Statement_INTL_Details"
           RstDet.Open SQL, cn, adOpenKeyset, adLockOptimistic
        For i = 1 To Me.ListView1.ListItems.Count
           RstDet.AddNew
                RstDet![SAID] = .Fields("SAID").Value
                RstDet![SAno] = .Fields("SAno").Value
                RstDet![Pax Name] = Me.ListView1.ListItems(i).SubItems(2)
                RstDet![Ticket No] = Me.ListView1.ListItems(i).SubItems(3)
           RstDet.Update
        Next i
    End If
         cn.CommitTrans
        If UCase(Me.Frame1.Caption) = "OVER-RIDE" Then
         MsgBox "SA International successfully updated...", vbInformation
        Else
         MsgBox "SA International successfully save...", vbInformation
        End If
       Rst.Close
     Set Rst = Nothing
     
     Dim Rpt As New RptSA_INTL_Show
     Dim AskPrint As Integer
     AskPrint = MsgBox("PLEASE INSERT PAPER AND CLICK OK TO START PRINTING...", vbOKCancel + vbExclamation)
      If AskPrint = vbNo Then: Exit Sub
        With Rpt
              .DataControl1.Connection = cn
              .DataControl1.Source = "SELECT * FROM qrytbl_Statement_INTL WHERE [SAno]='" & Me.txtSAno & "'"
              .Show 1
        End With
      'Call cmdNew_Click
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
MsgBox "An error occured while saving the SA " & Err.Description, vbInformation
End Sub

Function RemoveDetails(Param)
On Error GoTo FailSafe_Err
Dim Rst As New ADODB.Recordset
SQL = "DELETE * FROM tbl_Statement_INTL_Details WHERE [SAID]=" & Param
With cn
            .BeginTrans
                    .Execute SQL
            .CommitTrans
End With
Exit Function
FailSafe_Err:
cn.RollbackTrans
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
Private Sub cmdRecalc_Click()
If OK_2_Proceed Then
        Call Recalc
End If
End Sub

Function CompareValues(usrVal1, usrVal2) As Boolean

usrVal2 = IIf(CheckNull(usrVal2), 0, usrVal2)

If CDbl(usrVal1) >= CDbl(usrVal2) Then
        CompareValues = True
Else
        CompareValues = False
End If
End Function

Private Sub cmdRemovePax_Click()
If Me.ListView1.ListItems.Count > 0 Then
    Me.ListView1.ListItems.Remove lngIndex
End If
End Sub

Private Sub cmdSearchPO_Click()
frmPO_INTL_Find.Tag = "statement"
frmPO_INTL_Find.Show 1
End Sub

Private Sub Command1_Click()
If Me.ListView1.ListItems.Count > 0 Then
    Me.ListView1.ListItems.Remove lngIndex
End If
End Sub

Private Sub Form_Load()
Me.txtDate = RetDate(Now)
Me.txtIssuedBy = MDImain.StatusBar1.Panels(2).Text
Call FillAccountName
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


Function GetLastNumber() As String
Dim RsFnumber       As ADODB.Recordset
Dim SQL             As String
Dim Tmp             As String
Dim myTmpPos        As Integer

Set RsFnumber = New ADODB.Recordset
SQL = "SELECT [SAno] from tbl_Statement_INTL ORDER by [SAno] ASC"
RsFnumber.Open SQL, cn, adOpenKeyset, adLockOptimistic

With RsFnumber
        If .RecordCount > 0 Then
                .MoveLast
               Tmp = RsFnumber("SAno").Value
               myTmpPos = Int(ReturnFirst(Tmp)) - (Int(ReturnFirst(Tmp)) - Int(Return_1stDash(Tmp)))
               Tmp = Mid(Tmp, Return_1stDash(Tmp) + 5, myTmpPos - 1)
               GetLastNumber = "SI" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & AutoIncrement(Tmp)
        Else
               GetLastNumber = "SI" & WhichBranch.Fields(2).Value & Mid(Year(Now), 3, 2) & "-" & returnMon & returnDay & "000000000"
        End If
End With

End Function

Function Recalc() As Double
Dim TmpBasic(1 To 2)    As Double
Dim TmpDeduc(1 To 2)    As Double
Dim TmpEvat(1 To 2)     As Double
Dim TmpGrandTot(1 To 2) As Double
Dim PassCtr As Long
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
       
        
        
        
        Me.txtEvat(0) = RetCurrency(TmpEvat(1))
        Me.txtEvat(1) = RetCurrency(TmpEvat(2))
        
        
        TmpGrandTot(1) = CDbl(txtSubTot_Deduc(0)) + TmpEvat(1)
        TmpGrandTot(2) = CDbl(txtSubTot_Deduc(1)) + TmpEvat(2)
        
        
        txtSubTot_Deduc(0) = RetCurrency(TmpBasic(1) - TmpDeduc(1))
        txtSubTot_Deduc(1) = RetCurrency(TmpBasic(2) - TmpDeduc(2))
        
        
            Me.txtFare(0) = RetCurrency(Me.txtFare(0))
            Me.txtFare(1) = RetCurrency(Me.txtFare(1))
            Me.txtVUSA(0) = RetCurrency(Me.txtVUSA(0))
            Me.txtVUSA(1) = RetCurrency(Me.txtVUSA(1))
            Me.txtSwing(0) = RetCurrency(Me.txtSwing(0))
            Me.txtSwing(1) = RetCurrency(Me.txtSwing(1))
            Me.txtHotel(0) = RetCurrency(Me.txtHotel(0))
            Me.txtHotel(1) = RetCurrency(Me.txtHotel(1))
            Me.txtcar(0) = RetCurrency(Me.txtcar(0))
            Me.txtcar(1) = RetCurrency(Me.txtcar(1))
            Me.txttaxIns(0) = RetCurrency(Me.txttaxIns(0))
            Me.txttaxIns(1) = RetCurrency(Me.txttaxIns(1))
            Me.txtDomestic(0) = RetCurrency(Me.txtDomestic(0))
            Me.txtDomestic(1) = RetCurrency(Me.txtDomestic(1))
            Me.txtMiscName = RetCurrency(Me.txtMiscName)
            Me.txtMisc(0) = RetCurrency(Me.txtMisc(0))
            Me.txtMisc(1) = RetCurrency(Me.txtMisc(1))
            Me.txtPhilTravelTax(0) = RetCurrency(Me.txtPhilTravelTax(0))
            Me.txtPhilTravelTax(1) = RetCurrency(Me.txtPhilTravelTax(1))
        
        
        
        Me.txtGrandTot(0) = RetCurrency(TmpGrandTot(1))
        Me.txtGrandTot(1) = RetCurrency(TmpGrandTot(2))
        
        PassCtr = Me.ListView1.ListItems.Count
        Me.lblPassCtr = PassCtr
        Me.txtGrandTot_All(0) = RetCurrency(CDbl(TmpGrandTot(1)) * PassCtr)
        Me.txtGrandTot_All(1) = RetCurrency(CDbl(TmpGrandTot(2)) * PassCtr)
        
   
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
        
            Me.txtPONumber = IIf(Not IsNull(![Po Number]), ![Po Number], "")
            Me.txtDate = IIf(Not IsNull(![PO Date]), ![PO Date], "")
            Me.txtTo = IIf(Not IsNull(![Pay to]), ![Pay to], "")
            Me.txtRoute = IIf(Not IsNull(!Route), !Route, "")
            Me.txtRecordLoc = IIf(Not IsNull(![Record Locator]), ![Record Locator], "")
            Me.txtOthers = IIf(Not IsNull(!Others), !Others, "")
            Me.txtFOC = IIf(Not IsNull(![FOC no]), ![FOC no], "")
            Me.txtParticular = IIf(Not IsNull(!Particulars), !Particulars, "")
            
            Me.txtFare(0) = IIf(Not IsNull(![Fare Dollar]), RetCurrency(![Fare Dollar]), "0.00")
            Me.txtFare(1) = IIf(Not IsNull(![Fare Peso]), RetCurrency(![Fare Peso]), "0.00")
            Me.txtVUSA(0) = IIf(Not IsNull(![VUSA Dollar]), RetCurrency(![VUSA Dollar]), "0.00")
            Me.txtVUSA(1) = IIf(Not IsNull(![VUSA Peso]), RetCurrency(![VUSA Peso]), "0.00")
            Me.txtSwing(0) = IIf(Not IsNull(![Swing Around Dollar]), RetCurrency(![Swing Around Dollar]), "0.00")
            Me.txtSwing(1) = IIf(Not IsNull(![Swing Around Peso]), RetCurrency(![Swing Around Peso]), "0.00")
            Me.txtHotel(0) = IIf(Not IsNull(![Hotel Acco Dollar]), RetCurrency(![Hotel Acco Dollar]), "0.00")
            Me.txtHotel(1) = IIf(Not IsNull(![Hotel Acco Peso]), RetCurrency(![Hotel Acco Peso]), "0.00")
            Me.txtcar(0) = IIf(Not IsNull(![Car Rental Dollar]), RetCurrency(![Car Rental Dollar]), "0.00")
            Me.txtcar(1) = IIf(Not IsNull(![Car Rental Peso]), RetCurrency(![Car Rental Peso]), "0.00")
            Me.txttaxIns(0) = IIf(Not IsNull(![Tax Ins Dollar]), RetCurrency(![Tax Ins Dollar]), "0.00")
            Me.txttaxIns(1) = IIf(Not IsNull(![Tax Ins Peso]), RetCurrency(![Tax Ins Peso]), "0.00")
            Me.txtDomestic(0) = IIf(Not IsNull(![Domestic Ticket Dollar]), RetCurrency(![Domestic Ticket Dollar]), "0.00")
            Me.txtDomestic(1) = IIf(Not IsNull(![Domestic Ticket Peso]), RetCurrency(![Domestic Ticket Peso]), "0.00")
            Me.txtMiscName = IIf(Not IsNull(![Misc Name]), RetCurrency(![Misc Name]), "0.00")
            Me.txtMisc(0) = IIf(Not IsNull(![Misc Amount Dollar]), RetCurrency(![Misc Amount Dollar]), "0.00")
            Me.txtMisc(1) = IIf(Not IsNull(![Misc Amount Peso]), RetCurrency(![Misc Amount Peso]), "0.00")
            Me.txtPhilTravelTax(0) = IIf(Not IsNull(![Phil Travel Tax Dollar]), RetCurrency(![Phil Travel Tax Dollar]), "0.00")
            Me.txtPhilTravelTax(1) = IIf(Not IsNull(![Phil Travel Tax Peso]), RetCurrency(![Phil Travel Tax Peso]), "0.00")

'================================================================================================================================
'Load to txtPO for comparison later
'================================================================================================================================
            Me.txtFarePO(0) = IIf(Not IsNull(![Fare Dollar]), RetCurrency(![Fare Dollar]), "0.00")
            Me.txtFarePO(1) = IIf(Not IsNull(![Fare Peso]), RetCurrency(![Fare Peso]), "0.00")
            Me.txtVUSAPO(0) = IIf(Not IsNull(![VUSA Dollar]), RetCurrency(![VUSA Dollar]), "0.00")
            Me.txtVUSAPO(1) = IIf(Not IsNull(![VUSA Peso]), RetCurrency(![VUSA Peso]), "0.00")
            Me.txtSwingPO(0) = IIf(Not IsNull(![Swing Around Dollar]), RetCurrency(![Swing Around Dollar]), "0.00")
            Me.txtSwingPO(1) = IIf(Not IsNull(![Swing Around Peso]), RetCurrency(![Swing Around Peso]), "0.00")
            Me.txtHotelPO(0) = IIf(Not IsNull(![Hotel Acco Dollar]), RetCurrency(![Hotel Acco Dollar]), "0.00")
            Me.txtHotelPO(1) = IIf(Not IsNull(![Hotel Acco Peso]), RetCurrency(![Hotel Acco Peso]), "0.00")
            Me.txtcarPO(0) = IIf(Not IsNull(![Car Rental Dollar]), RetCurrency(![Car Rental Dollar]), "0.00")
            Me.txtcarPO(1) = IIf(Not IsNull(![Car Rental Peso]), RetCurrency(![Car Rental Peso]), "0.00")
            Me.txttaxInsPO(0) = IIf(Not IsNull(![Tax Ins Dollar]), RetCurrency(![Tax Ins Dollar]), "0.00")
            Me.txttaxInsPO(1) = IIf(Not IsNull(![Tax Ins Peso]), RetCurrency(![Tax Ins Peso]), "0.00")
            Me.txtDomesticPO(0) = IIf(Not IsNull(![Domestic Ticket Dollar]), RetCurrency(![Domestic Ticket Dollar]), "0.00")
            Me.txtDomesticPO(1) = IIf(Not IsNull(![Domestic Ticket Peso]), RetCurrency(![Domestic Ticket Peso]), "0.00")
            Me.txtMiscPO(0) = IIf(Not IsNull(![Misc Amount Dollar]), RetCurrency(![Misc Amount Dollar]), "0.00")
            Me.txtMiscPO(1) = IIf(Not IsNull(![Misc Amount Peso]), RetCurrency(![Misc Amount Peso]), "0.00")
            Me.txtPhilTravelTaxPO(0) = IIf(Not IsNull(![Phil Travel Tax Dollar]), RetCurrency(![Phil Travel Tax Dollar]), "0.00")
            Me.txtPhilTravelTaxPO(1) = IIf(Not IsNull(![Phil Travel Tax Peso]), RetCurrency(![Phil Travel Tax Peso]), "0.00")
'================================================================================================================================
'End Load
'================================================================================================================================
            
            Me.txtSubTot_Basic(0) = IIf(Not IsNull(![SubTot Basic Dollar]), RetCurrency(![SubTot Basic Dollar]), "0.00")
            Me.txtSubTot_Basic(1) = IIf(Not IsNull(![SubTot Basic Peso]), RetCurrency(![SubTot Basic Peso]), "0.00")
            Me.txtCommSwing = IIf(Not IsNull(![Swing Comm Percent]), RetCurrency(![Swing Comm Percent]), "0.00")
            Me.txtSwing_Less(0) = IIf(Not IsNull(![Swing Comm Dollar]), RetCurrency(![Swing Comm Dollar]), "0.00")
            Me.txtSwing_Less(1) = IIf(Not IsNull(![Swing Comm Peso]), RetCurrency(![Swing Comm Peso]), "0.00")
            Me.txtCommEvat = IIf(Not IsNull(![Evat Percent]), RetCurrency(![Evat Percent]), "0.00")
            Me.txtEvat(0) = IIf(Not IsNull(![Evat Dollar]), RetCurrency(![Evat Dollar]), "0.00")
            Me.txtEvat(1) = IIf(Not IsNull(![Evat Peso]), RetCurrency(![Evat Peso]), "0.00")
            Me.txtSubTot_Deduc(0) = IIf(Not IsNull(![SubTot Deduc Dollar]), RetCurrency(![SubTot Deduc Dollar]), "0.00")
            Me.txtSubTot_Deduc(1) = IIf(Not IsNull(![SubTot Deduc Peso]), RetCurrency(![SubTot Deduc Peso]), "0.00")
            Me.txtGrandTot(0) = IIf(Not IsNull(![Total Dollar]), RetCurrency(![Total Dollar]), "0.00")
            Me.txtGrandTot(1) = IIf(Not IsNull(![Total Peso]), RetCurrency(![Total Peso]), "0.00")
            Me.txtGrandTot_All(0) = IIf(Not IsNull(![Grand Total Dollar]), RetCurrency(![Grand Total Dollar]), "0.00")
            Me.txtGrandTot_All(1) = IIf(Not IsNull(![Grand Total Peso]), RetCurrency(![Grand Total Peso]), "0.00")
            Me.txtIssuedBy = IIf(Not IsNull(![Issued By]), RetCurrency(![Issued By]), "0.00")
       If .Fields("Posted").Value = "POSTED" Then
                Me.Combo1.ListIndex = 0
            Else
                Me.Combo1.ListIndex = 1
       End If
       Me.txtExchangeRate = IIf(Not IsNull(![Exchange Rate]), RetCurrency(![Exchange Rate]), "0.00")
         Me.Tag = IIf(Not IsNull(.Fields("PoID").Value), .Fields("PoID").Value, "")

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


Sub LoadValuesSA(Param)
Dim Rst                 As New ADODB.Recordset
Dim RsPODetails         As New ADODB.Recordset
Dim mySQL               As String
Dim ctr                 As Integer

  
SQL = "SELECT * from tbl_Statement_INTL WHERE [SAID]=" & Param & " ORDER by [SAno] ASC"

With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        
            Me.txtPONumber = IIf(Not IsNull(![Po Number]), ![Po Number], "")
            Me.txtSAno = IIf(Not IsNull(![SAno]), ![SAno], "")
            Me.txtDate = IIf(Not IsNull(![SA Date]), ![SA Date], "")
            Me.txtTo = IIf(Not IsNull(![Pay to]), ![Pay to], "")
            Me.txtRoute = IIf(Not IsNull(!Route), !Route, "")
            Me.txtRecordLoc = IIf(Not IsNull(![Record Locator]), ![Record Locator], "")
            Me.txtOthers = IIf(Not IsNull(!Others), !Others, "")
            Me.txtFOC = IIf(Not IsNull(![FOC no]), ![FOC no], "")
            Me.txtParticular = IIf(Not IsNull(!Particulars), !Particulars, "")
        
            
            Me.txtFare(0) = IIf(Not IsNull(![Fare Dollar]), RetCurrency(![Fare Dollar]), "0.00")
            Me.txtFare(1) = IIf(Not IsNull(![Fare Peso]), RetCurrency(![Fare Peso]), "0.00")
            Me.txtVUSA(0) = IIf(Not IsNull(![VUSA Dollar]), RetCurrency(![VUSA Dollar]), "0.00")
            Me.txtVUSA(1) = IIf(Not IsNull(![VUSA Peso]), RetCurrency(![VUSA Peso]), "0.00")
            Me.txtSwing(0) = IIf(Not IsNull(![Swing Around Dollar]), RetCurrency(![Swing Around Dollar]), "0.00")
            Me.txtSwing(1) = IIf(Not IsNull(![Swing Around Peso]), RetCurrency(![Swing Around Peso]), "0.00")
            Me.txtHotel(0) = IIf(Not IsNull(![Hotel Acco Dollar]), RetCurrency(![Hotel Acco Dollar]), "0.00")
            Me.txtHotel(1) = IIf(Not IsNull(![Hotel Acco Peso]), RetCurrency(![Hotel Acco Peso]), "0.00")
            Me.txtcar(0) = IIf(Not IsNull(![Car Rental Dollar]), RetCurrency(![Car Rental Dollar]), "0.00")
            Me.txtcar(1) = IIf(Not IsNull(![Car Rental Peso]), RetCurrency(![Car Rental Peso]), "0.00")
            Me.txttaxIns(0) = IIf(Not IsNull(![Tax Ins Dollar]), RetCurrency(![Tax Ins Dollar]), "0.00")
            Me.txttaxIns(1) = IIf(Not IsNull(![Tax Ins Peso]), RetCurrency(![Tax Ins Peso]), "0.00")
            Me.txtDomestic(0) = IIf(Not IsNull(![Domestic Ticket Dollar]), RetCurrency(![Domestic Ticket Dollar]), "0.00")
            Me.txtDomestic(1) = IIf(Not IsNull(![Domestic Ticket Peso]), RetCurrency(![Domestic Ticket Peso]), "0.00")
            Me.txtMiscName = IIf(Not IsNull(![Misc Name]), RetCurrency(![Misc Name]), "0.00")
            Me.txtMisc(0) = IIf(Not IsNull(![Misc Amount Dollar]), RetCurrency(![Misc Amount Dollar]), "0.00")
            Me.txtMisc(1) = IIf(Not IsNull(![Misc Amount Peso]), RetCurrency(![Misc Amount Peso]), "0.00")
            Me.txtPhilTravelTax(0) = IIf(Not IsNull(![Phil Travel Tax Dollar]), RetCurrency(![Phil Travel Tax Dollar]), "0.00")
            Me.txtPhilTravelTax(1) = IIf(Not IsNull(![Phil Travel Tax Peso]), RetCurrency(![Phil Travel Tax Peso]), "0.00")
            
            Me.txtSubTot_Basic(0) = IIf(Not IsNull(![SubTot Basic Dollar]), RetCurrency(![SubTot Basic Dollar]), "0.00")
            Me.txtSubTot_Basic(1) = IIf(Not IsNull(![SubTot Basic Peso]), RetCurrency(![SubTot Basic Peso]), "0.00")
            Me.txtCommSwing = IIf(Not IsNull(![Swing Comm Percent]), RetCurrency(![Swing Comm Percent]), "0.00")
            Me.txtSwing_Less(0) = IIf(Not IsNull(![Swing Comm Dollar]), RetCurrency(![Swing Comm Dollar]), "0.00")
            Me.txtSwing_Less(1) = IIf(Not IsNull(![Swing Comm Peso]), RetCurrency(![Swing Comm Peso]), "0.00")
            Me.txtCommEvat = IIf(Not IsNull(![Evat Percent]), RetCurrency(![Evat Percent]), "0.00")
            Me.txtEvat(0) = IIf(Not IsNull(![Evat Dollar]), RetCurrency(![Evat Dollar]), "0.00")
            Me.txtEvat(1) = IIf(Not IsNull(![Evat Peso]), RetCurrency(![Evat Peso]), "0.00")
            Me.txtSubTot_Deduc(0) = IIf(Not IsNull(![SubTot Deduc Dollar]), RetCurrency(![SubTot Deduc Dollar]), "0.00")
            Me.txtSubTot_Deduc(1) = IIf(Not IsNull(![SubTot Deduc Peso]), RetCurrency(![SubTot Deduc Peso]), "0.00")
            
            
            Me.txtGrandTot(0) = IIf(Not IsNull(![Total Dollar]), RetCurrency(![Total Dollar]), "0.00")
            Me.txtGrandTot(1) = IIf(Not IsNull(![Total Peso]), RetCurrency(![Total Peso]), "0.00")
            Me.txtGrandTot_All(0) = IIf(Not IsNull(![Grand Total Dollar]), RetCurrency(![Grand Total Dollar]), "0.00")
            Me.txtGrandTot_All(1) = IIf(Not IsNull(![Grand Total Peso]), RetCurrency(![Grand Total Peso]), "0.00")
            
            Me.txtIssuedBy = IIf(Not IsNull(![Issued By]), RetCurrency(![Issued By]), "0.00")
            
            
            
       If .Fields("Posted").Value = "POSTED" Then
                Me.Combo1.ListIndex = 0
            Else
                Me.Combo1.ListIndex = 1
       End If
       
        Me.txtExchangeRate = IIf(Not IsNull(![Exchange Rate]), RetCurrency(![Exchange Rate]), "0.00")
        Me.CboAccountName = IIf(Not IsNull(![AccountNo]), RetCurrency(![AccountNo]), "0.00")
        Me.txtAgencyName = IIf(Not IsNull(![AgencyName]), RetCurrency(![AgencyName]), "0.00")
         
         
         'Me.Tag = .Fields("PoID").Value
         
         

'//=======================================================================
'//pull out data from details and load it to list view
'//=======================================================================
mySQL = "SELECT * FROM tbl_Statement_INTL_Details WHERE [SAid]=" & Param
Me.ListView1.ListItems.Clear
         With RsPODetails
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then

                .MoveFirst
                    ctr = 0
                    On Error Resume Next
                    Do While Not .EOF
                    ctr = ctr + 1
                        ListView1.ListItems.Add , , .Fields("SAIDDetails").Value
                        ListView1.ListItems.Item(ctr).SubItems(1) = .Fields("SAID").Value
                        ListView1.ListItems.Item(ctr).SubItems(2) = .Fields("Pax Name").Value
                        ListView1.ListItems.Item(ctr).SubItems(3) = .Fields("Ticket No").Value
                        .MoveNext
                    Loop
               End If
        End With


        End If
End With

End Sub


Function OK_2_Proceed() As Boolean
        If CompareValues(Me.txtFare(0), Me.txtFarePO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount ", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtFare(0))
                Exit Function
        End If
        
        If CompareValues(Me.txtFare(1), Me.txtFarePO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtFare(1))
                Exit Function
        End If
        If CompareValues(Me.txtVUSA(0), Me.txtVUSAPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtVUSA(0))
                Exit Function
        End If
        If CompareValues(Me.txtVUSA(1), Me.txtVUSAPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtVUSA(1))
                Exit Function
        End If
        If CompareValues(Me.txtSwing(0), Me.txtSwingPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtSwing(0))
                Exit Function
        End If
        If CompareValues(Me.txtSwing(1), Me.txtSwingPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtSwing(1))
                Exit Function
        End If
        If CompareValues(Me.txtHotel(0), Me.txtHotelPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtHotel(0))
                Exit Function
        End If
        If CompareValues(Me.txtHotel(1), Me.txtHotelPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtHotel(1))
                Exit Function
        End If
        If CompareValues(Me.txtcar(0), Me.txtcarPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtFare(0))
                Exit Function
        End If
        If CompareValues(Me.txtcar(1), Me.txtcarPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtcar(1))
                Exit Function
        End If
        If CompareValues(Me.txttaxIns(0), Me.txttaxInsPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txttaxIns(0))
                Exit Function
        End If
        If CompareValues(Me.txttaxIns(1), Me.txttaxInsPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txttaxIns(1))
                Exit Function
        End If
        If CompareValues(Me.txtDomestic(0), Me.txtDomesticPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtDomestic(0))
                Exit Function
        End If
        If CompareValues(Me.txtDomestic(1), Me.txtDomesticPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtDomestic(1))
                Exit Function
        End If
        If CompareValues(Me.txtMisc(0), Me.txtMiscPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtMisc(0))
                Exit Function
        End If
        If CompareValues(Me.txtMisc(1), Me.txtMiscPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtMisc(1))
                Exit Function
        End If
        If CompareValues(Me.txtPhilTravelTax(0), Me.txtPhilTravelTaxPO(0)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtPhilTravelTax(0))
                Exit Function
        End If
        If CompareValues(Me.txtPhilTravelTax(1), Me.txtPhilTravelTaxPO(1)) Then
                OK_2_Proceed = True
        Else
                MsgBox "Invalid Amount", vbInformation
                OK_2_Proceed = False
                Call kulotHL(Me.txtPhilTravelTax(1))
                Exit Function
        End If

End Function

