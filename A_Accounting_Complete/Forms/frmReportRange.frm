VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmReportRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Report"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9225
   Icon            =   "frmReportRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   15
      ScaleHeight     =   6165
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Name"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   3345
         TabIndex        =   14
         Top             =   1590
         Width           =   5550
         Begin VB.ComboBox CboAccountName 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   315
            Width           =   5370
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   15
         ScaleHeight     =   855
         ScaleWidth      =   9255
         TabIndex        =   9
         Top             =   5355
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
                  Picture         =   "frmReportRange.frx":030A
                  Key             =   "IMG1"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":0FE4
                  Key             =   "IMG2"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":1E36
                  Key             =   "IMG3"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":2C88
                  Key             =   "IMG4"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":3562
                  Key             =   "IMG5"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":3E3C
                  Key             =   "IMG6"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":4716
                  Key             =   "IMG7"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":50E0
                  Key             =   "IMG8"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":59BA
                  Key             =   "IMG9"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":5CD4
                  Key             =   "IMG10"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":65AE
                  Key             =   "IMG11"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":6E88
                  Key             =   "IMG12"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":7762
                  Key             =   "IMG13"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":7A7C
                  Key             =   "IMG14"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":8356
                  Key             =   "IMG15"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":8C30
                  Key             =   "IMG16"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":950A
                  Key             =   "IMG17"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":9DE4
                  Key             =   "IMG18"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":A6BE
                  Key             =   "IMG19"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":AF98
                  Key             =   "IMG20"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":B872
                  Key             =   "IMG21"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":C14C
                  Key             =   "IMG22"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":CA26
                  Key             =   "IMG23"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":D300
                  Key             =   "IMG24"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":DBDA
                  Key             =   "IMG25"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":E4B4
                  Key             =   "IMG26"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":ED8E
                  Key             =   "IMG27"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":F668
                  Key             =   "IMG28"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":FF42
                  Key             =   "IMG29"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":1081C
                  Key             =   "IMG30"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":110D2
                  Key             =   "IMG31"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":119AC
                  Key             =   "IMG32"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":11DFE
                  Key             =   "IMG33"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":12250
                  Key             =   "IMG34"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":14A02
                  Key             =   "IMG35"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReportRange.frx":15C84
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin LVbuttons.LaVolpeButton cmdPrint 
            Height          =   480
            Left            =   6255
            TabIndex        =   11
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
            MICON           =   "frmReportRange.frx":18436
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
            TabIndex        =   12
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
            MICON           =   "frmReportRange.frx":18452
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
         Height          =   1740
         Left            =   240
         TabIndex        =   6
         Top             =   735
         Width           =   2895
         Begin VB.OptionButton OptPreview 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Preview"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptPrinter 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Send Directly to printer"
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         ForeColor       =   &H00FF0000&
         Height          =   720
         Left            =   3375
         TabIndex        =   1
         Top             =   810
         Width           =   5535
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   53870593
            CurrentDate     =   37497
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3600
            TabIndex        =   3
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   53870593
            CurrentDate     =   37497
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   3120
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label lblReportName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Financial Report"
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
         Left            =   0
         TabIndex        =   13
         Top             =   4845
         Width           =   9165
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
         TabIndex        =   10
         Top             =   210
         Width           =   4920
      End
   End
End
Attribute VB_Name = "frmReportRange"
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
Call Prn(Me.lblReportName)
End Sub

Sub Prn(Param)
Dim Rpt         As New RptFinancial
Dim Rpt1        As New RptRefund
Dim Rpt_        As Object
Dim SQL         As String
Dim Criteria    As String

If UCase(Param) = "VOUCHER" Then
Dim Rpt_Show As New RptVoucherShow
           SQL = "select * from QryVoucherRpt WHERE [Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
            With Rpt_Show
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
           
End If


If UCase(Param) = "PURCHASE ORDER" Then
Dim Rpt_ShowPO As New RptPO_Show
           SQL = "select * from qryPORpt WHERE [PO Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
            With Rpt_ShowPO
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
           
End If


If Param = "Purchase Order International" Then
Dim Rpt_ShowPo_intl As New RptPO_INTL_Show

           SQL = "SELECT * FROM qryTbl_PO_INTL WHERE [PO Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
            With Rpt_ShowPo_intl
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
           
End If



If UCase(Param) = "FINANCIAL" Then
           SQL = "select * from qryBankPassbook WHERE [Deposit Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
            With Rpt
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
           
End If
If UCase(Param) = "REFUND" Then
           SQL = "select * from qryRefund WHERE [Refund Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
            With Rpt1
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
           
End If


If UCase(Param) = "AR" Then
Dim myRptAR As New RptARDetailed
           SQL = "select * from qryAREport WHERE [Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "# AND [Account Name]='" & Me.CboAccountName & "' ORDER by [Date],[SNumber] asc "
            With myRptAR
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    .lblTotal = Format(frmReportRange.ReturnCustBal, "###,##0.00")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
           
End If

If UCase(Param) = "LIST OF SA" Then
Dim myRptListOfSA As New RptListOfSA
           SQL = "select * from qryListofSA WHERE [Date] between #" & _
           Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & _
           Format(Me.DTPicker2, "mm/dd/yyyy") & "#"
            With myRptListOfSA
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblFrom = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    .lblTO = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With

End If

If UCase(Param) = "STATEMENT OF ACCOUNTS" Then
Dim myRptSAofAccounts As New RptSAofAccounts
           SQL = "select * from qryStatementAsOf WHERE " & _
                 "[Account Name]='" & Me.CboAccountName & "'"
            With myRptSAofAccounts
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblAsOf = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
End If

If UCase(Param) = "AGING" Then
Dim myRptAging As New RptAging
           SQL = "select * from qryAging"
            With myRptAging
                    .DataControl1.Connection = cn
                    .DataControl1.Source = SQL
                    .lblAsOf = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                    If Me.OptPreview Then
                        .Show 1
                    Else
                        .PrintReport True
                    End If
            End With
End If



End Sub
Private Sub DTPicker1_Change()
'    If DTPicker1.Value > DTPicker2.Value Then
'        DTPicker2.Value = DTPicker1.Value
'    End If
End Sub

Private Sub DTPicker2_Change()
'    If DTPicker2.Value > DTPicker1.Value Then
'        DTPicker1.Value = DTPicker2.Value
'    End If
End Sub

Private Sub Form_Activate()
If UCase(Me.lblReportName) = "AR" Or UCase(Me.lblReportName) = "STATEMENT OF ACCOUNTS" Then
    Me.CboAccountName.Enabled = True
Else
    Me.CboAccountName.Enabled = False
End If

End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    Call FillAccountName
End Sub

Private Sub OptSelected_Click()

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

Function ReturnCustBal() As Double
Dim mySQL As String
Dim myRst As New ADODB.Recordset
Dim sDate As Date
Dim eDate As Date
Dim Tmp As Double


sDate = Format(Me.DTPicker1, "mm/dd/yyyy")
eDate = Format(Me.DTPicker2, "mm/dd/yyyy")

mySQL = "SELECT * FROM QryAR WHERE [AccountNo]='" & Me.CboAccountName & "' AND [Date] Between #" & sDate & "# And #" & eDate & "# AND [Paid]=False"

With myRst
        .Open mySQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        
        .MoveFirst
        Tmp = 0
            Do While Not .EOF
            Tmp = Tmp + .Fields("Balance").Value
            .MoveNext
            Loop
            ReturnCustBal = Tmp
        End If
        .Close
      Set myRst = Nothing
End With
End Function

