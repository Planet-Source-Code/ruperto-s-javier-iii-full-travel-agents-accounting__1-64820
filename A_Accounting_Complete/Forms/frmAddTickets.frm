VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAddTickets 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Ship / Airline"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   11925
   Begin VB.Frame Frame3 
      Height          =   810
      Left            =   8055
      TabIndex        =   28
      Top             =   7260
      Width           =   3780
      Begin VB.OptionButton OptWith 
         Caption         =   "Add Tickets With Purchase Order"
         Height          =   285
         Left            =   315
         TabIndex        =   30
         Top             =   465
         Width           =   3345
      End
      Begin VB.OptionButton OptWithOut 
         Caption         =   "Add Tickets With Out Purchase Order"
         Height          =   285
         Left            =   315
         TabIndex        =   29
         Top             =   165
         Value           =   -1  'True
         Width           =   3285
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   120
      TabIndex        =   22
      Top             =   6885
      Visible         =   0   'False
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   19
      Top             =   3465
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "TicketID"
         Caption         =   "TicketID"
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
         DataField       =   "AirlineID"
         Caption         =   "AirlineID"
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
         DataField       =   "TicketTypeID"
         Caption         =   "TicketTypeID"
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
         DataField       =   "AirlineName"
         Caption         =   "AirlineName"
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
         DataField       =   "Ticket Type"
         Caption         =   "Ticket Type"
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
         DataField       =   "Ticket No"
         Caption         =   "Ticket No"
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
         DataField       =   "Status"
         Caption         =   "Status"
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
      BeginProperty Column07 
         DataField       =   "Encoder"
         Caption         =   "Encoder"
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
      BeginProperty Column09 
         DataField       =   "Issued By"
         Caption         =   "Issued By"
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
         DataField       =   "Date Issued"
         Caption         =   "Date Issued"
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
         DataField       =   "PO"
         Caption         =   "PO"
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
         DataField       =   "Statement No"
         Caption         =   "Statement No"
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
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2340.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2145.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2445.166
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   11895
      TabIndex        =   13
      Top             =   0
      Width           =   11955
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Adding new data of Tickets"
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
         TabIndex        =   15
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Tickets"
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
         TabIndex        =   14
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.TextBox txtTotalCount 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4110
      TabIndex        =   12
      Text            =   "0"
      Top             =   7710
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Height          =   1020
      Left            =   135
      ScaleHeight     =   960
      ScaleWidth      =   11640
      TabIndex        =   6
      Top             =   2355
      Width           =   11700
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display All"
         Height          =   990
         Left            =   9690
         TabIndex        =   31
         Top             =   -15
         Width           =   1980
      End
      Begin VB.TextBox txtStart 
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
         Height          =   420
         Left            =   2655
         TabIndex        =   8
         Top             =   330
         Width           =   2295
      End
      Begin VB.TextBox txtEnd 
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
         Height          =   420
         Left            =   7215
         TabIndex        =   7
         Top             =   345
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Ticket No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   645
         TabIndex        =   10
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "End Ticket No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5250
         TabIndex        =   9
         Top             =   390
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Which Airline/Shipping Line and Ticket Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   1
      Top             =   1050
      Width           =   11700
      Begin VB.ComboBox cboTicketType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Text            =   "cboTicketType"
         Top             =   780
         Width           =   4680
      End
      Begin VB.ComboBox cboAirline 
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Text            =   "cboAirline"
         Top             =   780
         Width           =   4680
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Type"
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
         Height          =   255
         Left            =   5235
         TabIndex        =   4
         Top             =   435
         Width           =   4200
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Line / Airline"
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
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   4200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   180
      TabIndex        =   0
      Top             =   8055
      Width           =   11655
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3450
         TabIndex        =   32
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdUnsold 
         Caption         =   "Un-sold"
         Height          =   495
         Left            =   2340
         TabIndex        =   27
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   495
         Left            =   5445
         TabIndex        =   25
         Top             =   225
         Width           =   1530
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Find Ticket"
         Height          =   510
         Left            =   6990
         TabIndex        =   24
         Top             =   210
         Width           =   1290
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Range"
         Height          =   345
         Left            =   8325
         TabIndex        =   21
         Top             =   450
         Width           =   1425
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Individual"
         Height          =   345
         Left            =   8340
         TabIndex        =   20
         Top             =   150
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   1230
         TabIndex        =   18
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdAddSave 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   9990
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label lblRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Rec"
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
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   7275
      Width           =   4200
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3525
      TabIndex        =   23
      Top             =   7215
      Width           =   4200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Currently entered tickets :"
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
      Height          =   255
      Left            =   165
      TabIndex        =   11
      Top             =   7710
      Width           =   4200
   End
End
Attribute VB_Name = "frmAddTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim RsTickets As New ADODB.Recordset
Dim SQL As String
Dim OldBookMark
Dim RstAirline As New ADODB.Recordset
Dim Temp As String


Private Sub cboAirline_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM qryTickets WHERE [AirlineID]=" & FindAirlineID(Me.cboAirline)
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs
Call FillTicketType
End Sub

Private Sub cboTicketType_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM qryTickets WHERE [AirlineID]=" & FindAirlineID(Me.cboAirline) & " AND [Ticket Type]='" & Me.cboTicketType & "'"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs
Me.lblRec = "RECORD :" & Rs.AbsolutePosition & "/" & Rs.RecordCount
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
        Set Rs = New ADODB.Recordset
        SQL = "SELECT  * FROM qryTickets WHERE [AirlineID]=" & FindAirlineID(Me.cboAirline)
        Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
        Set Me.DataGrid1.DataSource = Rs
End If
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ErrExit

If Me.Tag = "proceed" Then: GoTo Cont
If Me.OptWithOut Then
   
        If Me.cmdAddSave.Caption = "Add" Then
            Me.cmdAddSave.Caption = "Save"
            
            Me.cboAirline.Enabled = True
            Me.cboTicketType.Enabled = True
            Me.txtStart.Enabled = True
            Me.txtEnd.Enabled = True
         
        Else
Cont:
            If CheckNull(Me.cboAirline) Then
                MsgBox "The Airline/Shipping Line Should Not be Blank"
                Exit Sub
            End If
            
            If CheckNull(Me.cboAirline) Then
                MsgBox "The Airline/Shipping Line Should Not be Blank"
                Exit Sub
            End If
            
            If CheckNull(Me.txtStart) Then
                MsgBox "The Starting Ticket No. Should Not be Blank"
                Exit Sub
            End If
            
            If CheckNull(Me.cboTicketType) Then
                MsgBox "Ticket Type should not br blank", vbCritical
                Exit Sub
            End If
            
            If Not IsNumeric(Me.txtStart) Then
                MsgBox "The Starting Ticket No. Should be a numeric value"
                Exit Sub
            End If
            
            If CheckNull(Me.txtEnd) Then
                MsgBox "The Ending Ticket No. Should Not be Blank"
                Exit Sub
            End If
            
            If Not IsNumeric(Me.txtEnd) Then
                MsgBox "The Ending Ticket No. Should be a numeric value"
                Exit Sub
            End If
        
            If CDbl(Me.txtStart) > CDbl(Me.txtEnd) Then
                MsgBox "Start ticket should be less than end ticket!"
                With Me.txtStart
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SetFocus
                End With
            Exit Sub
            End If
            
            Me.txtTotalCount = "0"
            Dim ctr As Long
            ctr = 0
            Dim i As Double
            Dim myMax As Double
            
            Me.ProgressBar1.Visible = True
            myMax = (CDbl(Me.txtEnd) - CDbl(Me.txtStart)) + 1
            Me.ProgressBar1.Min = 0
            Me.ProgressBar1.Max = 100
            Screen.MousePointer = vbHourglass
            lblStat.Caption = "Please wait.....saving..."
            Me.DataGrid1.Enabled = False
            
            For i = CDbl(Me.txtStart) To CDbl(Me.txtEnd) Step 1
            DoEvents
                If (Not (DupTicket(i))) And (Not (DupTicket("0" & i))) Then
                ctr = ctr + 1
                        With RsTickets
                        cn.BeginTrans
                                .AddNew
                                If Mid(Me.txtStart, 1, 1) = "0" Then
                                    .Fields(1).Value = "0" & CDbl(i)
                                Else
                                    .Fields(1).Value = CDbl(i)
                                End If
                                .Fields(2).Value = FindAirlineID(Me.cboAirline)
                                .Fields(3).Value = FindTicketTypeID(Me.cboTicketType)
                                .Fields(4).Value = "Un-Sold"
                                .Fields(5).Value = Format(Now, "mm/dd/yyyy")
                                .Fields("Encoder").Value = MDImain.StatusBar1.Panels(2).Text
                               If Me.Tag = "proceed" Then
                                    .Fields("PO").Value = frmAddTicketsSel.txtPONumber
                               End If
                                .Update
                                Me.ProgressBar1.Value = (ctr / myMax) * 100
                        cn.CommitTrans
                                Me.txtTotalCount = ctr
                        End With
                Else
                        MsgBox "Duplicate ticket Detected ===> 0" & i & " Press ok to skip...", vbInformation, "ELS"
                End If
            Next i
            Me.cmdAddSave.Caption = "Add"
            
            Me.cboAirline.Enabled = False
            Me.cboTicketType.Enabled = False
            Me.txtStart.Enabled = False
            Me.txtEnd.Enabled = False
            Rs.Requery
            
            Me.txtTotalCount = ctr
            If CDbl(ctr) > 0 Then
            Call UpDatePODetails(frmAddTicketsSel.txtPODetailsID)
                MsgBox "Tickets successfully added", vbInformation
            End If
            Me.ProgressBar1.Visible = False
            lblStat.Caption = ""
            Screen.MousePointer = vbDefault
            Me.DataGrid1.Enabled = True
        End If

Else
    'So the Tickets have PO
    If Me.OptWith Then
        frmPO_DomesticFind.Tag = "addtickets"
        frmPO_DomesticFind.Show 1
    End If
End If

Exit Sub

ErrExit:
cn.RollbackTrans
MsgBox "There was an error while trying to save the tickets please close the form and try again", vbCritical

End Sub
Function UpDatePODetails(Param)
On Error GoTo FailSafe_Err
Dim lngIndex As Long
SQL = "UPDATE tbl_PODetails_Domestic SET tbl_PODetails_Domestic.Encoded = True " & _
    "WHERE (((tbl_PODetails_Domestic.POID_Details)=" & Param & ") AND ((tbl_PODetails_Domestic.Encoded)=False));"
    
    cn.BeginTrans
        cn.Execute SQL
    cn.CommitTrans
    lngIndex = CLng(frmAddTicketsSel.txtSelected)
    frmAddTicketsSel.ListView1.ListItems.Item(lngIndex).SubItems(7) = "Done"
Exit Function
FailSafe_Err:
cn.RollbackTrans
MsgBox "There was an error while saving this ticket", vbCritical
End Function

Private Sub cmdAddSave_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cmdCancel_Click()
Me.cmdAddSave.Caption = "Add"
            Me.cboAirline.Enabled = False
            Me.cboTicketType.Enabled = False
            Me.txtStart.Enabled = False
            Me.txtEnd.Enabled = False
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim ask As Integer
Dim ctr As Long
Dim i As Double
Dim myMax As Double
    
ask = MsgBox("Are you sure you want to delete?", vbCritical + vbYesNo)
If ask = vbYes Then

If Me.Option1 Then
ctr = 1
    SQL = "DELETE * FROM tbl_Tickets WHERE [TicketID]=" & Me.DataGrid1.Columns(0).Text
    cn.Execute SQL
End If

If Me.Option2 Then

If CheckNull(Me.txtStart) Then: MsgBox "Start ticket should not be blank": Exit Sub
If CheckNull(Me.txtEnd) Then: MsgBox "end ticket should not be blank": Exit Sub

If CDbl(Me.txtStart) > CDbl(Me.txtEnd) Then
    MsgBox "Start ticket should not be greater than end ticket", vbCritical
    Exit Sub
End If
    ctr = 0
    Screen.MousePointer = vbHourglass
    
    Me.ProgressBar1.Visible = True
    myMax = (CDbl(Me.txtEnd) - CDbl(Me.txtStart)) + 1
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = 100
    Me.lblStat = "Please Wait... Deleting..."
    Me.DataGrid1.Enabled = False
    
    For i = CDbl(Me.txtStart) To CDbl(Me.txtEnd) Step 1
    DoEvents
    
        If ((DupTicket(i))) Or ((DupTicket("0" & i))) Then
        ctr = ctr + 1
                With RsTickets
                cn.BeginTrans
                
                    SQL = "DELETE * FROM tbl_Tickets WHERE [Ticket No]='" & i & "'"
                    cn.Execute SQL
                    
                    SQL = "DELETE * FROM tbl_Tickets WHERE [Ticket No]='0" & i & "'"
                    cn.Execute SQL
                    
                    
                    Me.ProgressBar1.Value = (ctr / myMax) * 100
                    cn.CommitTrans
                    Me.txtTotalCount = ctr
                End With
        ' Else
        '        MsgBox "Not found!!! ticket  ===> 0" & i & " Press ok to skip...", vbInformation, "ELS"
         End If
    Next i
End If

        OldBookMark = Rs.Bookmark
        Screen.MousePointer = vbDefault
        Me.lblStat = ""
        Rs.Requery
        Me.ProgressBar1.Value = 100
        MsgBox ctr & " Ticket(s) deleted", vbInformation
        Me.DataGrid1.Enabled = True
        Rs.Bookmark = OldBookMark
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFilter_Click()

Me.txtSearch.Enabled = True
Call kulotHL(Me.txtSearch)
End Sub

Private Sub cmdUnsold_Click()
On Error GoTo FailSafe_Err
Dim myTicketID  As Long
Dim myTicketNo  As String
Dim myBookMark  As Long

myTicketID = IIf(Not (IsNull(Me.DataGrid1.Columns(0).Text)), CLng(Me.DataGrid1.Columns(0).Text), -1)


If myTicketID <= 0 Then
        MsgBox "Please select a ticket to update", vbInformation
        Exit Sub
End If

myTicketNo = Me.DataGrid1.Columns(5).Text

If CheckInSA(myTicketNo) = "found" Then
    Dim ask As Integer
     ask = MsgBox("Mark this ticket as Un-Sold and remove the issued ticket?", vbInformation + vbYesNo)
     If ask = vbNo Then: Exit Sub
     cn.BeginTrans
            cn.Execute "DELETE * FROM tbl_StatementDetail WHERE [Ticket No]='" & myTicketNo & "'"
     cn.CommitTrans
     MsgBox "The Ticket [" & myTicketNo & "] was removed from sales", vbInformation, "Info"
End If


SQL = "UPDATE tbl_Tickets SET tbl_Tickets.Status = 'Un-Sold', tbl_Tickets.[Issued By] = '', tbl_Tickets.[Date Issued] = null WHERE (((tbl_Tickets.TicketID)=" & myTicketID & "));"
cn.BeginTrans
        cn.Execute SQL
cn.CommitTrans

Call Search_Ticket(myTicketNo)
Exit Sub
FailSafe_Err:
cn.RollbackTrans
MsgBox "There was an error while trying to unsold", vbInformation
End Sub

Function CheckInSA(usrTicket) As String
Dim RstCheck As New ADODB.Recordset
SQL = "SELECT * FROM tbl_StatementDetail WHERE [Ticket No]='" & usrTicket & "'"

With RstCheck
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
               Call MsgBox("The ticket youre trying to un-sold was already sold to " & .Fields("Name").Value, vbInformation, "Warning!!!")
               CheckInSA = "found"
            Else
               CheckInSA = "notfound"
        End If
        .Close
      Set RstCheck = Nothing
End With
End Function

Private Sub Form_Load()
Call FillAirline
Call FillTicketType

Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM qryTickets WHERE [AirlineID]=" & FindAirlineID(Me.cboAirline)
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs

Set RsTickets = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_Tickets"
RsTickets.Open SQL, cn, adOpenKeyset, adLockOptimistic

End Sub




Sub DisplayTicketType(ByVal Param As String)
Set RsTicketType = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_TicketType WHERE [AirlineShippingLine]='" & UCase(Param) & "' ORDER BY [Ticket Type] ASC"
RsTicketType.Open SQL, cn, adOpenKeyset, adLockOptimistic
'Set Me.DataGrid3.DataSource = RsTicketType
End Sub
'====


Function FindAirlineID(Param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindAirlineID = .Fields(0).Value
          Else
              FindAirlineID = -1
        End If
     .Close
   Set Rst = Nothing
End With
End Function


Function FindTicketTypeID(Param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_TicketType WHERE [Ticket Type]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindTicketTypeID = .Fields(0).Value
          Else
              FindTicketTypeID = -1
        End If
     .Close
   Set Rst = Nothing
End With
End Function


Function DupTicket(Ticket) As Boolean
Dim Tmp As ADODB.Recordset

Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_Tickets WHERE [Ticket No]='" & UCase(Ticket) & "'"    ' AND [AirlineID]=" & FindAirlineID(Me.cboAirline)
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            DupTicket = True
        Else
            DupTicket = False
        End If
      .Close
    Set Tmp = Nothing
End With

End Function


Sub FillAirline()
SQL = "SELECT  * FROM tbl_Airline WHERE [AirlineName]<>'NONE' ORDER by [AirlineName] ASC "
        Me.cboAirline.Clear
        
        
With RstAirline
        If .State = 1 Then
            .Close
        End If
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboAirline.AddItem .Fields(1).Value
                .MoveNext
            Loop
         Me.cboAirline.ListIndex = 0
        End If
        
       '.Close
     'Set RstAirline = Nothing
End With
End Sub


Sub FillTicketType()
Dim Rst As New ADODB.Recordset
Dim RsTemps As ADODB.Recordset
Dim ReturnRec As String
Dim STRSQL As String

Set RsTemps = New ADODB.Recordset
STRSQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Me.cboAirline & "'"

RsTemps.Open STRSQL, cn, adOpenKeyset, adLockOptimistic
If RsTemps.RecordCount > 0 Then
    ReturnRec = RsTemps.Fields(2).Value
End If
RsTemps.Close
Set RsTemps = Nothing

SQL = "SELECT  * FROM tbl_TicketType WHERE [AirlineShippingLine]='" & UCase(ReturnRec) & "' ORDER BY [Ticket Type] ASC"
        Me.cboTicketType.Clear
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboTicketType.AddItem .Fields(1).Value
                .MoveNext
            Loop
        End If
        .Close
       ' Me.cboTicketType.ListIndex = 0
     Set Rst = Nothing

End With

End Sub


Sub Search_Ticket(usrSearch)
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM qryTickets WHERE [Ticket No] like '" & usrSearch & "%'"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs
Me.lblRec = "RECORD :" & Rs.AbsolutePosition & "/" & Rs.RecordCount
End Sub

Private Sub Option1_Click()
Me.txtStart = ""
Me.txtEnd = ""
End Sub

Private Sub Option2_Click()
Me.txtStart.Enabled = True
Me.txtEnd.Enabled = True
End Sub


Private Sub txtSearch_Change()
Call Search_Ticket(Me.txtSearch)
End Sub

Private Sub txtStart_Change()
Me.txtEnd = Me.txtStart
End Sub
