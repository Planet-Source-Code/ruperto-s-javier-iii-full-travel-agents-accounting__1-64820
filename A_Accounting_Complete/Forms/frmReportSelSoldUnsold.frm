VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmReportSelSoldUnsold 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sold UnSold Tickets Report"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9225
   Icon            =   "frmReportSelSoldUnsold.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   15
      ScaleHeight     =   5535
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company Option"
         ForeColor       =   &H00FF0000&
         Height          =   1635
         Left            =   3480
         TabIndex        =   19
         Top             =   2460
         Width           =   5535
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select Shipping/Airline Below"
            ForeColor       =   &H000000FF&
            Height          =   1140
            Left            =   2685
            TabIndex        =   22
            Top             =   300
            Width           =   2700
            Begin VB.ComboBox Combo1 
               Enabled         =   0   'False
               Height          =   315
               Left            =   150
               TabIndex        =   23
               Text            =   "Combo1"
               Top             =   540
               Width           =   2430
            End
         End
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Shipping/Airline"
            Height          =   270
            Left            =   255
            TabIndex        =   21
            Top             =   330
            Value           =   -1  'True
            Width           =   1980
         End
         Begin VB.OptionButton OptIndividual 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Individual Shipping/Airline"
            Height          =   270
            Left            =   255
            TabIndex        =   20
            Top             =   900
            Width           =   2325
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Criteria"
         ForeColor       =   &H00FF0000&
         Height          =   1635
         Left            =   225
         TabIndex        =   16
         Top             =   2460
         Width           =   2895
         Begin VB.OptionButton OptUnsold 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Un-Sold Tickets"
            Height          =   270
            Left            =   390
            TabIndex        =   18
            Top             =   900
            Width           =   1980
         End
         Begin VB.OptionButton OptSold 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sold Tickets"
            Height          =   270
            Left            =   390
            TabIndex        =   17
            Top             =   330
            Value           =   -1  'True
            Width           =   1980
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   9255
         TabIndex        =   11
         Top             =   4545
         Width           =   9255
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
                  Picture         =   "frmReportSelSoldUnsold.frx":030A
                  Key             =   "IMG1"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":0FE4
                  Key             =   "IMG2"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":1E36
                  Key             =   "IMG3"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":2C88
                  Key             =   "IMG4"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":3562
                  Key             =   "IMG5"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":3E3C
                  Key             =   "IMG6"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":4716
                  Key             =   "IMG7"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":50E0
                  Key             =   "IMG8"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":59BA
                  Key             =   "IMG9"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":5CD4
                  Key             =   "IMG10"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":65AE
                  Key             =   "IMG11"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":6E88
                  Key             =   "IMG12"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":7762
                  Key             =   "IMG13"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":7A7C
                  Key             =   "IMG14"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":8356
                  Key             =   "IMG15"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":8C30
                  Key             =   "IMG16"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":950A
                  Key             =   "IMG17"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":9DE4
                  Key             =   "IMG18"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":A6BE
                  Key             =   "IMG19"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":AF98
                  Key             =   "IMG20"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":B872
                  Key             =   "IMG21"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":C14C
                  Key             =   "IMG22"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":CA26
                  Key             =   "IMG23"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":D300
                  Key             =   "IMG24"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":DBDA
                  Key             =   "IMG25"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":E4B4
                  Key             =   "IMG26"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":ED8E
                  Key             =   "IMG27"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":F668
                  Key             =   "IMG28"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":FF42
                  Key             =   "IMG29"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":1081C
                  Key             =   "IMG30"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":110D2
                  Key             =   "IMG31"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":119AC
                  Key             =   "IMG32"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":11DFE
                  Key             =   "IMG33"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":12250
                  Key             =   "IMG34"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":14A02
                  Key             =   "IMG35"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportSelSoldUnsold.frx":15C84
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin LVbuttons.LaVolpeButton cmdPrint 
            Height          =   480
            Left            =   6255
            TabIndex        =   13
            Top             =   210
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   847
            BTYPE           =   3
            TX              =   "Print"
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
            MICON           =   "frmReportSelSoldUnsold.frx":18436
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
            Left            =   7665
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
            MICON           =   "frmReportSelSoldUnsold.frx":18452
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   240
         TabIndex        =   8
         Top             =   735
         Width           =   2895
         Begin VB.OptionButton OptPreview 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Preview"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptPrinter 
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
         Left            =   3480
         TabIndex        =   1
         Top             =   735
         Width           =   5535
         Begin VB.OptionButton OptSelected 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Selected Date"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton OptStart 
            BackColor       =   &H00FFFFFF&
            Caption         =   "From Start to Present"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   840
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   53477377
            CurrentDate     =   37497
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   2880
            TabIndex        =   5
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   53477377
            CurrentDate     =   37497
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Label lblReportName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sold-Unsold Tickets Report"
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
         Left            =   45
         TabIndex        =   15
         Top             =   4080
         Width           =   9120
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reports Generator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
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
Attribute VB_Name = "frmReportSelSoldUnsold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Call Prn
End Sub


Sub Prn()
Dim Rpt As New RptSoldUnsoldTickets
Dim SQL As String
Dim Criteria As String

If Me.OptSold Then
    Criteria = "Sold"
Else
    Criteria = "Un-Sold"
End If

If Me.OptStart Then
        If Me.OptAll Then
           SQL = "select * from QrySoldUnSoldTkts WHERE [Status]='" & Criteria & "'"
        Else
           SQL = "select * from QrySoldUnSoldTkts WHERE [Status]='" & Criteria & "' AND [AirlineName]='" & Me.Combo1 & "'"
        End If
Else
        If Me.OptAll Then
           SQL = "select * from QrySoldUnSoldTkts WHERE [Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "# AND [Status]='" & Criteria & "'"

        Else
           SQL = "select * from QrySoldUnSoldTkts WHERE [Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "# AND [Status]='" & Criteria & "' AND [AirlineName]='" & Me.Combo1 & "'"
        End If
End If
       
            With Rpt
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    
                    .Label27 = "TOTAL NO. OF " & Criteria & " TICKETS :"
                If Me.OptPreview Then
                    .Show 1
                Else
                    .PrintReport True
                End If
            End With

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
Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    Call FillCombo
End Sub

Private Sub OptAll_Click()
Me.Combo1.Enabled = False
End Sub

Private Sub OptIndividual_Click()
Me.Combo1.Enabled = True
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
        Me.Combo1.ListIndex = 0
    End If
End With
End Sub

