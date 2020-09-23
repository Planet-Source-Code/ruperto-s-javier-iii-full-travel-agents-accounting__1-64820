VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmReportSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12465
   Icon            =   "frmReportSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12465
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Left            =   15
      ScaleHeight     =   7800
      ScaleWidth      =   12420
      TabIndex        =   0
      Top             =   0
      Width           =   12420
      Begin VB.PictureBox Picture3 
         Height          =   3225
         Left            =   75
         ScaleHeight     =   3165
         ScaleWidth      =   12195
         TabIndex        =   16
         Top             =   2790
         Width           =   12255
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<"
            Height          =   450
            Left            =   8565
            TabIndex        =   28
            Top             =   1605
            Width           =   990
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   ">"
            Height          =   450
            Left            =   8565
            TabIndex        =   27
            Top             =   1095
            Width           =   990
         End
         Begin VB.ListBox List2 
            Height          =   2400
            Left            =   9705
            TabIndex        =   26
            Top             =   555
            Width           =   2460
         End
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   5910
            TabIndex        =   25
            Top             =   585
            Width           =   2460
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Company Option"
            ForeColor       =   &H00FF0000&
            Height          =   1635
            Left            =   45
            TabIndex        =   20
            Top             =   105
            Width           =   5535
            Begin VB.OptionButton OptIndividual 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Individual Shipping/Airline"
               Height          =   270
               Left            =   255
               TabIndex        =   24
               Top             =   900
               Value           =   -1  'True
               Width           =   2325
            End
            Begin VB.OptionButton OptAll 
               BackColor       =   &H00FFFFFF&
               Caption         =   "All Shipping/Airline"
               Enabled         =   0   'False
               Height          =   270
               Left            =   270
               TabIndex        =   23
               Top             =   330
               Width           =   1980
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Select Shipping/Airline Below"
               ForeColor       =   &H000000FF&
               Height          =   1140
               Left            =   2685
               TabIndex        =   21
               Top             =   300
               Width           =   2700
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   22
                  Text            =   "Combo1"
                  Top             =   540
                  Width           =   2430
               End
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Criteria"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   1185
            Left            =   30
            TabIndex        =   17
            Top             =   1770
            Width           =   5535
            Begin VB.OptionButton OptWOEvat 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Exclude EVAT computation for this report(Sales/Company Report Only)"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   660
               Width           =   5370
            End
            Begin VB.OptionButton optWithEvat 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Include EVAT computation for this report (Sales/Company Report Only)"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   18
               Top             =   255
               Value           =   -1  'True
               Width           =   5370
            End
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "Selected Ticket Type"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   9705
            TabIndex        =   30
            Top             =   195
            Width           =   2445
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "Ticket Type"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   5895
            TabIndex        =   29
            Top             =   210
            Width           =   2460
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   45
         ScaleHeight     =   855
         ScaleWidth      =   12360
         TabIndex        =   11
         Top             =   6810
         Width           =   12360
         Begin MSComctlLib.ImageList SmallImages 
            Left            =   2595
            Top             =   60
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
                  Picture         =   "frmReportSel.frx":030A
                  Key             =   "IMG1"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":0FE4
                  Key             =   "IMG2"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":1E36
                  Key             =   "IMG3"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":2C88
                  Key             =   "IMG4"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":3562
                  Key             =   "IMG5"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":3E3C
                  Key             =   "IMG6"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":4716
                  Key             =   "IMG7"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":50E0
                  Key             =   "IMG8"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":59BA
                  Key             =   "IMG9"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":5CD4
                  Key             =   "IMG10"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":65AE
                  Key             =   "IMG11"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":6E88
                  Key             =   "IMG12"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":7762
                  Key             =   "IMG13"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":7A7C
                  Key             =   "IMG14"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":8356
                  Key             =   "IMG15"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":8C30
                  Key             =   "IMG16"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":950A
                  Key             =   "IMG17"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":9DE4
                  Key             =   "IMG18"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":A6BE
                  Key             =   "IMG19"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":AF98
                  Key             =   "IMG20"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":B872
                  Key             =   "IMG21"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":C14C
                  Key             =   "IMG22"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":CA26
                  Key             =   "IMG23"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":D300
                  Key             =   "IMG24"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":DBDA
                  Key             =   "IMG25"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":E4B4
                  Key             =   "IMG26"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":ED8E
                  Key             =   "IMG27"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":F668
                  Key             =   "IMG28"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":FF42
                  Key             =   "IMG29"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":1081C
                  Key             =   "IMG30"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":110D2
                  Key             =   "IMG31"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":119AC
                  Key             =   "IMG32"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":11DFE
                  Key             =   "IMG33"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":12250
                  Key             =   "IMG34"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":14A02
                  Key             =   "IMG35"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSel.frx":15C84
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin LVbuttons.LaVolpeButton cmdPrint 
            Height          =   480
            Left            =   4980
            TabIndex        =   13
            Top             =   195
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Print to Company"
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
            MICON           =   "frmReportSel.frx":18436
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
         Begin LVbuttons.LaVolpeButton cmdExit 
            Height          =   480
            Left            =   10815
            TabIndex        =   14
            Top             =   210
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
            MICON           =   "frmReportSel.frx":18452
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
         Begin LVbuttons.LaVolpeButton LaVolpeButton1 
            Height          =   480
            Left            =   7530
            TabIndex        =   31
            Top             =   195
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Print Own Copy"
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
            MICON           =   "frmReportSel.frx":1846E
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   150
         TabIndex        =   8
         Top             =   960
         Width           =   2895
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Preview"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Send Directly to printer"
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   3060
         TabIndex        =   1
         Top             =   960
         Width           =   9225
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Selected Date"
            Height          =   255
            Left            =   1215
            TabIndex        =   3
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "From Start to Present"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1230
            TabIndex        =   2
            Top             =   360
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   5640
            TabIndex        =   4
            Top             =   360
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   556
            _Version        =   393216
            Format          =   53673985
            CurrentDate     =   37497
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   5625
            TabIndex        =   5
            Top             =   840
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   556
            _Version        =   393216
            Format          =   53673985
            CurrentDate     =   37497
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   5130
            TabIndex        =   7
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            Height          =   255
            Left            =   4890
            TabIndex        =   6
            Top             =   375
            Width           =   615
         End
      End
      Begin VB.Image Image1 
         Height          =   2085
         Left            =   2085
         Picture         =   "frmReportSel.frx":1848A
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   8055
      End
      Begin VB.Label lblReportName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   135
         TabIndex        =   15
         Top             =   6150
         Width           =   12165
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reports Generator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   225
         TabIndex        =   12
         Top             =   210
         Width           =   4920
      End
   End
End
Attribute VB_Name = "frmReportSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAirline As ADODB.Recordset
Dim SQL As String
Dim TempRecord As String
'Const MySearchStr = "SELECT tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID, tbl_Statement.Date," & _
'      "tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name, tbl_StatementDetail.Void AS nVoid," & _
'      "tbl_StatementTickets.Route,qSum_MISC.misc, IIf([nVoid]=False,[misc],0) AS nMisc, IIf([nVoid]=False,[Gross],0) AS nGross, IIf([nVoid]=False," & _
'      "[Company Commission],0) AS [nCompany Commission], tbl_StatementTickets.VAT AS xVAT," & _
'      "IIf([nVoid]=False,IIf(IsNull([xVat]),0,[xVat]),0) AS nVAT, IIf([nVoid]=False,[SumOfInsurance],0) AS nSumOfInsurance," & _
'      "tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF, IIf([nVoid]=False,[SumOfASF],0) AS nSumOfASF," & _
'      "IIf([nVoid]=False,[Terminal Fee],0) AS [nTerminal Fee], IIf([nVoid]=False,[SumOfTerminal Fee],0) AS [nSumOfTerminal Fee]," & _
'      "IIf([nVoid]=False,[Meals],0) AS nMeals, IIf([nVoid]=False,Format([nGross]*([nCompany Commission])/100,'Standard'),0) AS Comm," & _
'      "IIf([nVoid]=False,[Comm]*([nVat]/100),0) AS qVAT, tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineID, tbl_Airline.AirlineName," & _
'      "IIf([nVoid]=False,[nSumOfInsurance],0) AS TotIns, IIf([nVoid]=False,([nGross]+[TotIns]+[nSumOfASF]+[nSumOfTerminal Fee]+[nMeals])*0.1,0) AS qEvat," & _
'      "IIf([nVoid]=False,[nGross]-[Comm]+[TotIns]+[nSumOfASF]+[nSumOfTerminal Fee]+[nMeals]+[qvat],0) AS Subtotal, " & _
'      " IIf([nVoid]=False,IIf([AirlineID]=36,Abs(Int(([Subtotal]+[qEvat])*-1)),[Subtotal]+[qEvat]),0) AS CashDue, IIf([nvoid]=False,[Fare],0) AS nFare, [nFare]-IIf(IsNull([Cashdue]),0,[CashDue]) AS Profit FROM tbl_Statement INNER JOIN " & _
'      "(((((tbl_Airline RIGHT JOIN tbl_StatementDetail ON tbl_Airline.AirlineID = tbl_StatementDetail.Airline)" & _
'      " LEFT JOIN q1 ON tbl_StatementDetail.StatementDetails = q1.StatementDetails) LEFT JOIN qSum_ASF ON " & _
'      "tbl_StatementDetail.StatementDetails = qSum_ASF.StatementDetails) LEFT JOIN qSum_TF ON " & _
'      "tbl_StatementDetail.StatementDetails = qSum_TF.StatementDetails) INNER JOIN tbl_StatementTickets ON " & _
'      "tbl_StatementDetail.StatementDetails = tbl_StatementTickets.StatementDetails) ON " & _
'      "tbl_Statement.TransID = tbl_StatementDetail.TransID GROUP BY tbl_StatementDetail.StatementDetails," & _
'      " tbl_StatementDetail.TransID, tbl_Statement.Date, tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name," & _
'      " tbl_StatementDetail.Void, tbl_StatementTickets.Route, tbl_StatementTickets.VAT, tbl_StatementTickets.Insurance," & _
'      " tbl_StatementTickets.ASF, tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineID, tbl_Airline.AirlineName," & _
'      " tbl_StatementDetail.Fare, tbl_StatementDetail.Gross, tbl_StatementTickets.[Company Commission], q1.SumOfInsurance, qSum_ASF.SumOfASF," & _
'      " tbl_StatementTickets.[Terminal Fee], qSum_TF.[SumOfTerminal Fee], tbl_StatementTickets.Meals "


Const MySearchStr = "SELECT tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID, tbl_Statement.Date, " & _
        "tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name, tbl_StatementDetail.Void AS nVoid, " & _
        "tbl_StatementTickets.Route, qSum_MISC.misc, IIf([nVoid]=False,[misc],0) AS nMisc, IIf([nVoid]=False,[Gross],0) AS nGross, IIf([nVoid]=False," & _
        "[Company Commission],0) AS [nCompany Commission], tbl_StatementTickets.VAT AS xVAT," & _
        "IIf([nVoid]=False,IIf(IsNull([xVat]),0,[xVat]),0) AS nVAT, IIf([nVoid]=False,[SumOfInsurance],0) AS nSumOfInsurance," & _
        "tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF, IIf([nVoid]=False,[SumOfASF],0) AS nSumOfASF," & _
        "IIf([nVoid]=False,[Terminal Fee],0) AS [nTerminal Fee], IIf([nVoid]=False,[SumOfTerminal Fee],0) AS [nSumOfTerminal Fee]," & _
        "IIf([nVoid]=False,[Meals],0) AS nMeals, IIf([nVoid]=False,Format([nGross]*([nCompany Commission])/100,'Standard'),0) AS Comm," & _
        "IIf([nVoid]=False,[Comm]*([nVat]/100),0) AS qVAT, tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineID, tbl_Airline.AirlineName," & _
        "IIf([nVoid]=False,[nSumOfInsurance],0) AS TotIns, IIf([nVoid]=False,([nGross]+[TotIns]+[nSumOfASF]+[nSumOfTerminal Fee]+[nMeals]+[nMisc])*0.12,0) AS qEvat," & _
        "IIf([nVoid]=False,[nGross]-[Comm]+[TotIns]+[nSumOfASF]+[nSumOfTerminal Fee]+[nMeals]+[qvat]+[nMisc],0) AS Subtotal," & _
        "IIf([nVoid]=False,IIf([AirlineID]=36,Abs(Int(([Subtotal]+[qEvat])*-1)),[Subtotal]+[qEvat]),0) AS CashDue, IIf([nvoid]=False,[Fare],0) AS nFare, [nFare]-IIf(IsNull([Cashdue]),0,[CashDue]) AS Profit FROM tbl_Statement INNER JOIN " & _
        "((((((tbl_Airline RIGHT JOIN tbl_StatementDetail ON tbl_Airline.AirlineID = tbl_StatementDetail.Airline)" & _
        " LEFT JOIN q1 ON tbl_StatementDetail.StatementDetails = q1.StatementDetails) LEFT JOIN qSum_ASF ON " & _
        "tbl_StatementDetail.StatementDetails = qSum_ASF.StatementDetails) LEFT JOIN qSum_TF ON " & _
        "tbl_StatementDetail.StatementDetails = qSum_TF.StatementDetails) LEFT JOIN qSum_MISC ON " & _
        "tbl_StatementDetail.StatementDetails = qSum_MISC.StatementDetails) INNER JOIN tbl_StatementTickets ON " & _
        "tbl_StatementDetail.StatementDetails = tbl_StatementTickets.StatementDetails) ON tbl_Statement.TransID = tbl_StatementDetail.TransID " & _
        "GROUP BY tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID, tbl_Statement.Date, tbl_StatementDetail.[Ticket No], " & _
        "tbl_StatementDetail.Name , tbl_StatementDetail.Void, tbl_StatementTickets.Route, qSum_MISC.Misc, " & _
        "tbl_StatementTickets.VAT , tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF, " & _
        "tbl_StatementDetail.[Ticket Type] , tbl_Airline.AirlineID, tbl_Airline.AirlineName, " & _
        "tbl_StatementDetail.Fare , tbl_StatementDetail.Gross, tbl_StatementTickets.[Company Commission], " & _
        "q1.SumOfInsurance , qSum_ASF.SumOfASF, tbl_StatementTickets.[Terminal Fee], qSum_TF.[SumOfTerminal Fee], tbl_StatementTickets.Meals "

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdInsert_Click()
If Len(Me.List1.List(Me.List1.ListIndex)) > 0 Then
            Me.List2.AddItem Me.List1.List(Me.List1.ListIndex)
            Me.List1.RemoveItem (Me.List1.ListIndex)
End If
End Sub

Private Sub cmdRemove_Click()
If Len(Me.List2.List(Me.List2.ListIndex)) > 0 Then
            Me.List1.AddItem Me.List2.List(Me.List2.ListIndex)
            Me.List2.RemoveItem (Me.List2.ListIndex)
End If
End Sub
Private Sub cmdPrint_Click()
Call Prn(Me.lblReportName, 1)
End Sub


Sub Prn(param, btn)
Dim Rpt As New RptBankAcc
Dim Rpt1 As New RptSalesReport
Dim Rpt2 As New RptReportCompany
Dim Rpt3 As New RptBankAccPALCC
Dim Rpt4 As New RptSaleCompany
Dim Rpt5 As New RptCancelled
Dim Rpt6 As New RptSaleCompanyEvat
Dim myRptProfit As New RptProfit
Dim SQL As String
Dim sDate, eDate
Dim MyCriteria As String


'Me.Hide

Select Case param
Case "Bank Deposits Report"
'====================================================================================
'For bank deposits
        If Me.Option1 Then
            SQL = "select * from qryRptBankAccounts"
        End If
        If Me.Option2 Then
            SQL = "select * from qryRptBankAccounts WHERE [Date] between #" & Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
        End If
        
            With Rpt
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                If Me.Option5 Then
                    .Show
                Else
                    .PrintReport True
                End If
            End With
'====================================================================================
Case "Detailed Sales Report"
'forDetailed sales
        If Me.Option1 Then
            SQL = "select * from qrySalesRpt"
        End If
        If Me.Option2 Then
            SQL = "select * from qrySalesRpt WHERE [Date] between #" & Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
        End If

        With Rpt1
                .DataControl1.Connection = cn
                .DataControl1.Source = SQL
                If Me.Option5 Then
                    .Show 1
                Else
                    .PrintReport True
                End If
        End With

'====================================================================================
Case "Sales Report"
        If Me.Option1 Then
            SQL = "select * from qrySalesRpt"
        End If
        If Me.Option2 Then
            SQL = "select * from qrySalesRpt WHERE [Date] between #" & Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
        End If

        With Rpt2
                .DataControl1.Connection = cn
                .DataControl1.Source = SQL
                If Me.Option5 Then
                    .Show 1
                Else
                    .PrintReport True
                End If
        End With
'====================================================================================
Case "Company Profit Report"
        
        If Me.Option2 Then
                                sDate = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                                eDate = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                                
                                MyCriteria = "HAVING ((Date) Between #" & sDate & "# And #" & eDate & "#)" & _
                                             " ORDER BY [AirlineName], [Ticket No]"
                                SQL = MySearchStr & MyCriteria
            
        End If

        With myRptProfit
                .DataControl1.Connection = cn
                .DataControl1.Source = SQL
                If Me.Option5 Then
                    .Show
                Else
                    .PrintReport True
                End If
        End With
'====================================================================================
Case "Company Sales Report"
Dim i As Integer
Dim TicketType(1 To 3) As String 'Declare Dynamic Array
Dim TicketTypeFlag As Boolean
                                Dim tmp
                                Dim myTmpSQL As String

            If Me.OptIndividual Then
            
                 If Me.List2.ListCount > 0 Then
                 
                                sDate = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                                eDate = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                                
                                'MyCriteria = "HAVING (((Date) Between #" & sDate & "# And #" & eDate & "#)" & _
                                '             " AND (([Ticket Type])='" & Me.List2.List(0) & "' Or" & _
                                '             " ([Ticket Type])='" & Me.List2.List(1) & "')" & _
                                '             " AND (([AirlineName])= '" & Me.Combo1 & "' ))" & _
                                '             " OR ((([Ticket Type])='" & Me.List2.List(2) & "'" & _
                                '             " Or ([Ticket Type])='" & Me.List2.List(3) & "')) " & _
                                '             "ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                
                            If Me.List2.ListCount = 1 Then
                                MyCriteria = " where [AirlineName] ='" & Me.Combo1 & "'" & _
                                             " AND [Date] Between #" & sDate & "# AND #" & eDate & "#" & _
                                             " AND [Ticket Type]='" & Me.List2.List(0) & "'" & _
                                             " ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                SQL = "SELECT * FROM qrySalesRptEVAT " & MyCriteria
                            End If
                             
                             If Me.List2.ListCount = 2 Then
                             
                           
                                Call SpoolReport
                                
                                tmp = kulotRead(App.Path & "\Settings.txt")
                                Select Case CDbl(tmp)
                                Case 1
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool01"
                                Case 2
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool02"
                                Case 3
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool03"
                                Case 4
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool04"
                                End Select
                                MyCriteria = myTmpSQL & _
                                             " WHERE [Ticket Type]='" & Me.List2.List(0) & "'" & _
                                             " OR [Ticket Type]= '" & Me.List2.List(1) & "'" & _
                                             " ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                SQL = MyCriteria
                          
                             End If
                             
                             If Me.List2.ListCount = 3 Then
                                Call SpoolReport
                                
                                tmp = kulotRead(App.Path & "\Settings.txt")
                                Select Case CDbl(tmp)
                                Case 1
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool01"
                                Case 2
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool02"
                                Case 3
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool03"
                                Case 4
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool04"
                                End Select
                                MyCriteria = myTmpSQL & _
                                             " WHERE [Ticket Type]='" & Me.List2.List(0) & "'" & _
                                             " OR [Ticket Type]= '" & Me.List2.List(1) & "'" & _
                                             " OR [Ticket Type]= '" & Me.List2.List(2) & "'" & _
                                             " ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                SQL = MyCriteria
                             
                             End If
                                
                                
                            '==============================================
                                
                            With Rpt6
                            If btn = 1 Then
                                    .Field3.Visible = False
                            Else
                                    .Field3.Visible = True
                            End If
                                  
                                  
                                    .DataControl1.Connection = cn
                                    .DataControl1.Source = SQL
                                    
                                    If Me.Option5 Then
                                        'Unload Me
                                        .Show
                                    Else
                                        .PrintReport True
                                  
                                    End If
                            End With

                            '==============================================
                  End If
            End If
            

        
Case "Bank Deposits PAL Credit Card"
        If Me.Option1 Then
            SQL = "select * from qryRptBankAccountsPalCC"
        End If
        If Me.Option2 Then
            SQL = "select * from qryRptBankAccountsPalCC WHERE [Date] between #" & Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
        End If

        With Rpt3
                .DataControl1.Connection = cn
                .DataControl1.Source = SQL
                If Me.Option5 Then
                    .Show
                Else
                    .PrintReport True
                End If
        End With
'======================================================
Case "Cancelled Report"
        If Me.Option1 Then
            SQL = "select * from QryCancelled"
        End If
        If Me.Option2 Then
            SQL = "select * from QryCancelled WHERE [Date] between #" & Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
        End If

        With Rpt5
                .DataControl1.Connection = cn
                .DataControl1.Source = SQL
                .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")

                If Me.Option5 Then
                    .Show
                Else
                    .PrintReport True
                End If
        End With
End Select

End Sub


Private Sub Combo1_Click()
Call FillList
End Sub

Private Sub DTPicker1_Change()
    If DTPicker1.Value > DTPicker2.Value Then
        DTPicker2.Value = DTPicker1.Value
    End If
End Sub

Private Sub DTPicker2_Change()
    If DTPicker2.Value > DTPicker1.Value Then
        DTPicker1.Value = DTPicker2.Value
    End If
End Sub

Private Sub Form_Activate()

    If Me.lblReportName = "Company Sales Report" Then
        Me.Frame1.Enabled = True
    Else
        Me.Frame1.Enabled = False
    End If
    
    Me.List2.Clear
    
    Call FillCombo
    Call FillList
    
    If Me.lblReportName = "Company Profit Report" Then
        Me.Picture3.Visible = False
    End If
End Sub

Private Sub Form_Load()


    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    
    
    
    
End Sub


Sub FillCombo()
Me.Combo1.Clear
Set RsAirline = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]<>'NONE'"

With RsAirline
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
            Me.Combo1.AddItem .Fields(1).Value
            .MoveNext
        Loop
        Me.Combo1.ListIndex = 0
    End If
End With
End Sub


Sub FillList()
Dim tmp As ADODB.Recordset
Dim tmpAirline As New ADODB.Recordset
Dim SQL As String

Set tmpAirline = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Me.Combo1 & "'"
tmpAirline.Open SQL, cn, adOpenKeyset, adLockOptimistic
If tmpAirline.RecordCount > 0 Then
    TempRecord = tmpAirline.Fields(2).Value
End If


Set tmp = New ADODB.Recordset

SQL = "SELECT * FROM tbl_TicketType WHERE [AirlineShippingLine]='" & TempRecord & "' ORDER BY [Ticket Type] ASC"
Me.List1.Clear
With tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
        Me.List1.AddItem .Fields(1).Value
        .MoveNext
        Loop
        Me.List1.ListIndex = 0
    End If
End With

End Sub

Private Sub LaVolpeButton1_Click()
Call Prn(Me.lblReportName, 2)
End Sub

Private Sub OptAll_Click()
Me.Combo1.Enabled = False
End Sub

Private Sub OptIndividual_Click()
Me.Combo1.Enabled = True
End Sub

Private Sub Option1_Click()
Me.DTPicker1.Enabled = False
Me.DTPicker2.Enabled = False
End Sub

Private Sub Option2_Click()
'Me.DTPicker1.Enabled = True
'Me.DTPicker2.Enabled = True

End Sub


Sub SpoolReport()
'On Error GoTo FailSafe_Error
Dim mySQL               As String
Dim recs                As Long
Dim tmp
Dim myInsertSQL         As String
Dim myInsertSQL_Head    As String

tmp = kulotRead(App.Path & "\Settings.txt")




mySQL = "(StatementDetails, TransID, [Date], [Ticket No]," & _
                 "Name, nVoid, Route, misc, nMisc, nGross, [nCompany Commission]," & _
                 "xVAT, nVAT, nSumOfInsurance, Insurance, ASF, nSumOfASF," & _
                 "[nTerminal Fee], [nSumOfTerminal Fee], nMeals," & _
                 "Comm, qVAT, [Ticket Type], AirlineID, AirlineName,"

mySQL = mySQL & "TotIns, qEvat, Subtotal, CashDue, nFare, Profit, Fare, Gross," & _
                "[Company Commission], SumOfInsurance, SumOfASF, [Terminal Fee], " & _
                "[SumOfTerminal Fee], Meals)"

mySQL = mySQL & "SELECT tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID," & _
                "tbl_Statement.Date, tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name," & _
                "tbl_StatementDetail.Void AS nVoid, tbl_StatementTickets.Route, qSum_MISC.misc," & _
                " IIf([nVoid]=False,[misc],0) AS nMisc, IIf([nVoid]=False,[Gross],0) AS nGross," & _
                " IIf([nVoid]=False,[Company Commission],0) AS [nCompany Commission], " & _
                "tbl_StatementTickets.VAT AS xVAT,"
      
mySQL = mySQL & "IIf([nVoid]=False,IIf(IsNull([xVat]),0,[xVat]),0) AS nVAT," & _
                " IIf([nVoid]=False,[SumOfInsurance],0) AS nSumOfInsurance," & _
                " tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF," & _
                " IIf([nVoid]=False,[SumOfASF],0) AS nSumOfASF, IIf([nVoid]=False," & _
                "[Terminal Fee],0) AS [nTerminal Fee], IIf([nVoid]=False,[SumOfTerminal Fee],0)" & _
                " AS [nSumOfTerminal Fee], IIf([nVoid]=False,[Meals],0) AS nMeals, " & _
                "IIf([nVoid]=False,Format([nGross]*([nCompany Commission])/100,'Standard'),0) AS Comm,"
      
mySQL = mySQL & "IIf([nVoid]=False,[Comm]*([nVat]/100),0) AS qVAT, tbl_StatementDetail.[Ticket Type]," & _
                " tbl_Airline.AirlineID, tbl_Airline.AirlineName, IIf([nVoid]=False,[nSumOfInsurance],0)" & _
                " AS TotIns, IIf([nVoid]=False,([nGross]+[TotIns]+[nSumOfASF]+[nSumOfTerminal Fee]" & _
                "+[nMeals]+[nMisc])*0.12,0) AS qEvat, IIf([nVoid]=False,[nGross]-[Comm]+[TotIns]" & _
                "+[nSumOfASF]+[nSumOfTerminal Fee]+[nMeals]+[qvat]+[nMisc],0) AS Subtotal, " & _
                "IIf([nVoid]=False,IIf([AirlineID]=36,Abs(Int(([Subtotal]+[qEvat])*-1))," & _
                "[Subtotal]+[qEvat]),0) AS CashDue, IIf([nvoid]=False,[Fare],0) AS nFare, " & _
                "[nFare]-IIf(IsNull([Cashdue]),0,[CashDue]) AS Profit, tbl_StatementDetail.Fare," & _
                "tbl_StatementDetail.Gross, tbl_StatementTickets.[Company Commission]," & _
                "q1.SumOfInsurance, qSum_ASF.SumOfASF, tbl_StatementTickets.[Terminal Fee], " & _
                "qSum_TF.[SumOfTerminal Fee], tbl_StatementTickets.Meals"
mySQL = mySQL & " FROM tbl_Statement INNER JOIN ((((((tbl_Airline RIGHT JOIN tbl_StatementDetail" & _
                " ON tbl_Airline.AirlineID=tbl_StatementDetail.Airline) LEFT JOIN q1 " & _
                " ON tbl_StatementDetail.StatementDetails=q1.StatementDetails) " & _
                " LEFT JOIN qSum_ASF ON tbl_StatementDetail.StatementDetails=qSum_ASF.StatementDetails) " & _
                " LEFT JOIN qSum_TF ON tbl_StatementDetail.StatementDetails=qSum_TF.StatementDetails) " & _
                " LEFT JOIN qSum_MISC ON tbl_StatementDetail.StatementDetails=qSum_MISC.StatementDetails)" & _
                " INNER JOIN tbl_StatementTickets ON tbl_StatementDetail.StatementDetails " & _
                "=tbl_StatementTickets.StatementDetails) ON tbl_Statement.TransID=tbl_StatementDetail.TransID" & _
                " GROUP BY tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID, " & _
                "tbl_Statement.Date, tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name, " & _
                "tbl_StatementDetail.Void, tbl_StatementTickets.Route, qSum_MISC.misc, " & _
                "tbl_StatementTickets.VAT, tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF," & _
                " tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineID, tbl_Airline.AirlineName, " & _
                "tbl_StatementDetail.Fare, tbl_StatementDetail.Gross, tbl_StatementTickets.[Company Commission], q1.SumOfInsurance," & _
                " qSum_ASF.SumOfASF, tbl_StatementTickets.[Terminal Fee], qSum_TF.[SumOfTerminal Fee], tbl_StatementTickets.Meals " & _
                " HAVING (((tbl_Statement.Date) Between #" & Format(Me.DTPicker1, "MM/DD/YYYY") & "# And #" & Format(Me.DTPicker2, "MM/DD/YYYY") & "#) AND ((tbl_Airline.AirlineName)='" & Me.Combo1 & "'))" & _
                " ORDER BY tbl_StatementDetail.[Ticket No], tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineName;"



Select Case CDbl(tmp)
Case 1
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool01"
        mySQL = myInsertSQL_Head & mySQL
Case 2
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool02"
        mySQL = myInsertSQL_Head & mySQL

Case 3
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool03"
        mySQL = myInsertSQL_Head & mySQL

Case 4
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool04"
        mySQL = myInsertSQL_Head & mySQL

End Select

cn.BeginTrans
        'Remove first existing Records
         Select Case CDbl(tmp)
            Case 1
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool01"
            Case 2
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool02"
            Case 3
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool03"
            Case 4
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool04"
         End Select
        

        'Now append to spool table
        cn.Execute mySQL, recs, adExecuteNoRecords
cn.CommitTrans
Exit Sub

FailSafe_Error:
    cn.RollbackTrans
End Sub
