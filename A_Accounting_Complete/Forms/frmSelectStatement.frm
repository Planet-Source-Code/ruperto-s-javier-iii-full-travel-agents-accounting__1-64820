VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelectStatement 
   Caption         =   "Select Statement"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   Icon            =   "frmSelectStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearchNT 
      Caption         =   "Search By Name/Ticket #"
      Height          =   600
      Left            =   8925
      TabIndex        =   23
      Top             =   5895
      Width           =   2940
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   705
      Left            =   7200
      TabIndex        =   22
      Top             =   5130
      Width           =   1635
   End
   Begin VB.PictureBox Picture4 
      Height          =   1305
      Left            =   2460
      ScaleHeight     =   1245
      ScaleWidth      =   6375
      TabIndex        =   14
      Top             =   4575
      Width           =   6435
      Begin VB.PictureBox Picture5 
         Enabled         =   0   'False
         Height          =   825
         Left            =   -15
         ScaleHeight     =   765
         ScaleWidth      =   6345
         TabIndex        =   17
         Top             =   420
         Width           =   6405
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   2175
            TabIndex        =   20
            Top             =   45
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   582
            _Version        =   393216
            Format          =   58654721
            CurrentDate     =   38677
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   2175
            TabIndex        =   21
            Top             =   405
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   582
            _Version        =   393216
            Format          =   58654721
            CurrentDate     =   38677
         End
         Begin VB.Label Label5 
            Caption         =   "End Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   555
            TabIndex        =   19
            Top             =   450
            Width           =   1785
         End
         Begin VB.Label Label4 
            Caption         =   "Start Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   540
            TabIndex        =   18
            Top             =   30
            Width           =   1785
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Selected Date"
         Height          =   300
         Left            =   3495
         TabIndex        =   16
         Top             =   90
         Width           =   2310
      End
      Begin VB.OptionButton Option1 
         Caption         =   "From Start to Finish"
         Height          =   300
         Left            =   1410
         TabIndex        =   15
         Top             =   105
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1305
      Left            =   8910
      ScaleHeight     =   1245
      ScaleWidth      =   2895
      TabIndex        =   11
      Top             =   4575
      Width           =   2955
      Begin VB.Label lblSATotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   780
         TabIndex        =   13
         Top             =   105
         Width           =   2040
      End
      Begin VB.Label Label3 
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         TabIndex        =   12
         Top             =   105
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1275
      Left            =   105
      ScaleHeight     =   1215
      ScaleWidth      =   2265
      TabIndex        =   8
      Top             =   4590
      Width           =   2325
      Begin VB.Label lblStat 
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1200
         TabIndex        =   10
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Record :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   105
         TabIndex        =   9
         Top             =   90
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   -30
      ScaleHeight     =   915
      ScaleWidth      =   11940
      TabIndex        =   5
      Top             =   -30
      Width           =   12000
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Statement"
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
         Left            =   990
         TabIndex        =   7
         Top             =   75
         Width           =   3210
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Find Statement for individual viewing"
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
         Left            =   1005
         TabIndex        =   6
         Top             =   570
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmSelectStatement.frx":0442
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   435
      Left            =   8910
      TabIndex        =   4
      Top             =   6570
      Width           =   1470
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   10395
      TabIndex        =   3
      Top             =   6585
      Width           =   1470
   End
   Begin VB.TextBox txtSearch 
      Height          =   405
      Left            =   3675
      TabIndex        =   2
      Top             =   6585
      Width           =   5160
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3600
      Left            =   120
      TabIndex        =   0
      Top             =   945
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   6350
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "TransID"
         Caption         =   "TransID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "sNumber"
         Caption         =   "Statement Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Airline"
         Caption         =   "Airline"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "AccountNo"
         Caption         =   "AccountNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AgencyName"
         Caption         =   "AgencyName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Date"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Total Amount"
         Caption         =   "Total Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """PHP""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Airline"
         Caption         =   "Airline"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Branch Number"
         Caption         =   "Branch Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Credit Card Activated"
         Caption         =   "Credit Card Activated"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Paid"
         Caption         =   "Paid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Void"
         Caption         =   "Void"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Refund"
         Caption         =   "Refund"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Balance"
         Caption         =   "Balance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Down"
         Caption         =   "Down"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "AccountID"
         Caption         =   "AccountID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2880
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Type Statement No. here :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   630
      TabIndex        =   1
      Top             =   6615
      Width           =   2835
   End
End
Attribute VB_Name = "frmSelectStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstDisplay              As ADODB.Recordset
Dim SQL                     As String

Function ReturnSumSA() As Double
Dim RstSum As New ADODB.Recordset
Dim sDate
Dim eDate

        sDate = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
        eDate = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
If Me.Option1 Then
  'SQL = "SELECT * FROM qrySumSA"
    SQL = "SELECT Sum(tbl_Statement.[Total Amount]) AS [SumOfTotal Amount] FROM tbl_Statement;"
Else
    SQL = "SELECT * FROM qrySumSA WHERE [Date] between #" & sDate & "# AND #" & eDate & "#"
End If
With RstSum

    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
            ReturnSumSA = .Fields(0).Value
            Else
            ReturnSumSA = 0
    End If
    .Close
  Set RstSum = Nothing
End With


End Function
Sub ShowStatement(ByVal strSearch As String)
'On Error GoTo FailSafe_Error
Dim sDate
Dim eDate

' SQL = "between #" & Format(Me.DTPicker1.Value, "mm/dd/yyyy") & "# AND #" & Format(Me.DTPicker2, "mm/dd/yyyy") & "#"

If Me.Option1 Then
        If Len(strSearch) > 0 Then
                 SQL = "SELECT * FROM tbl_Statement WHERE [sNumber]='" & strSearch & "' ORDER by [AgencyName] ASC"
        Else
                 SQL = "SELECT * FROM tbl_Statement ORDER by [AgencyName] ASC"
        End If
End If
        
If Me.Option2 Then
        
        sDate = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
        eDate = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
        
        If Len(strSearch) > 0 Then
                 SQL = "SELECT * FROM tbl_Statement WHERE [Date] between #" & sDate & "# AND #" & eDate & "# AND [sNumber]='" & strSearch & "' ORDER by [AgencyName] ASC"
        Else
                 SQL = "SELECT * FROM tbl_Statement WHERE [Date] between #" & sDate & "# AND #" & eDate & "# ORDER by [AgencyName] ASC"
        End If
End If
        
        
                 Set RstDisplay = New ADODB.Recordset
                 With RstDisplay
                       .Open SQL, cn, adOpenKeyset, adLockOptimistic
                         If .RecordCount > 0 Then
                            Set Me.DataGrid1.DataSource = RstDisplay
                            Me.lblStat = .AbsolutePosition & "/" & .RecordCount
                            Me.lblSATotal = Format(ReturnSumSA(), "###,##0.00")
                         Else
                            Set Me.DataGrid1.DataSource = Nothing
                                Me.DataGrid1.Refresh
                                Me.lblSATotal = "0.00"
                                Me.lblStat = "0/0"
                        End If
                End With
        
     Exit Sub
FailSafe_Error:
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrExit
Dim Rpt As New RptStatement


Me.txtSearch = Me.DataGrid1.Columns(1).Text
If Me.Tag = "over_ride" Then
    frmStatement.txtNo = Me.DataGrid1.Columns(1).Text
    frmUserVerify.Show 1
Else
If Check_If_Printed(Me.txtSearch) Then
    Dim ask As Integer
          ask = MsgBox("This Statement was already printed!!! Over-Ride?", vbInformation + vbYesNo)
         If ask = vbNo Then: Exit Sub
         frmUserVerifyRefund.Tag = "Report Domestic Over-ride"
         frmUserVerifyRefund.txtSearch = Me.txtSearch
         frmUserVerifyRefund.Show 1
Else

     With Rpt
          .DataControl1.Connection = cn
          .DataControl1.Source = "SELECT * FROM qryStatement WHERE [sNumber]='" & Me.txtSearch & "' AND [Printed]=False ORDER by [Ticket No]"
          .Show 1
     End With
End If
End If
     Set Rpt = Nothing
Exit Sub
ErrExit:
MsgBox "Cannot find this statement! or this was not saved!", vbInformation
End Sub

Function Check_If_Printed(param) As Boolean
Dim Rst As New ADODB.Recordset
Dim mySQL As String
'SD106-032403147-1
mySQL = "SELECT * FROM qryStatement WHERE [sNumber]='" & param & "' AND [Printed]=TRUE ORDER by [Ticket No]"
With Rst
        .Open mySQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Check_If_Printed = True
              Else
                Check_If_Printed = False
           End If
        .Close
        Set Rst = Nothing
End With
End Function

Private Sub cmdRefresh_Click()
If CheckNull(Me.txtSearch) Then
    Call ShowStatement("")
Else
    Call ShowStatement(Me.txtSearch)
End If

End Sub

Private Sub cmdSearchNT_Click()
frmSearchName.Tag = "over_ride"
frmSearchName.Show 1
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'On Error GoTo FailSafe_Error
'Me.txtSearch = Me.DataGrid1.Columns(1).Text
'FailSafe_Error:
End Sub

Private Sub Form_Load()
Call ShowStatement("")
 
 DTPicker1.Value = Now
 DTPicker2.Value = Now
    
End Sub


Private Sub Option1_Click()
Me.Picture5.Enabled = False
End Sub

Private Sub Option2_Click()
Me.Picture5.Enabled = True
End Sub

Private Sub txtSearch_Change()
Call ShowStatement(Me.txtSearch)
End Sub

Private Sub txtSearch_GotFocus()
Call kulotHL(Me.txtSearch)
End Sub
