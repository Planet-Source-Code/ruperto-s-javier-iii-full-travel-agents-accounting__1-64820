VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearchName 
   Caption         =   "Search by Name "
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   Icon            =   "frmSearchName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   720
      Left            =   105
      ScaleHeight     =   660
      ScaleWidth      =   7170
      TabIndex        =   5
      Top             =   4185
      Width           =   7230
      Begin VB.OptionButton Option2 
         Caption         =   "Search By Ticket #"
         Height          =   315
         Left            =   3465
         TabIndex        =   7
         Top             =   150
         Width           =   2925
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Search By Name"
         Height          =   300
         Left            =   930
         TabIndex        =   6
         Top             =   150
         Value           =   -1  'True
         Width           =   1890
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Left            =   105
      ScaleHeight     =   585
      ScaleWidth      =   11730
      TabIndex        =   3
      Top             =   75
      Width           =   11790
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2025
         TabIndex        =   0
         Text            =   "type your search here...."
         Top             =   90
         Width           =   5265
      End
      Begin VB.Label Label1 
         Caption         =   "Search Here :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   105
         Width           =   1920
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   765
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5953
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "sNumber"
         Caption         =   "Statement #"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Name"
         Caption         =   "Name"
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
         DataField       =   "Total Amount"
         Caption         =   "Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2835.213
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2459.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1244.976
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Done"
      Height          =   480
      Left            =   9810
      TabIndex        =   1
      Top             =   4215
      Width           =   2085
   End
End
Attribute VB_Name = "frmSearchName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCust              As New ADODB.Recordset
Dim SQL                 As String


Private Sub cmdExit_Click()
If Me.Tag <> "over_ride" Then
    frmCashier.OptSearchSBS.Value = True
End If
Unload Me
End Sub



Private Sub DataGrid1_Click()
On Error GoTo FailSafe_Error
   If Me.Tag <> "over_ride" Then
            frmCashier.Text1 = Me.DataGrid1.Columns(0).Text
        If Me.Option1 Then
            Me.Text1 = Me.DataGrid1.Columns(3).Text
        Else
            Me.Text1 = Me.DataGrid1.Columns(2).Text
        End If
   Else
            frmSelectStatement.txtSearch = Me.DataGrid1.Columns(0).Text
   End If
Exit Sub
FailSafe_Error:
   If Me.Tag <> "over_ride" Then
            frmCashier.Text1 = ""
   End If
            Me.Text1 = "type your search here...."
End Sub

Private Sub Form_Load()
Call kulotHL(Me.Text1)
Call Disp_Cust("")
End Sub


Sub Disp_Cust(ByVal Criteria As String)
If Me.Option1 Then
        If Len(Criteria) < 1 Then
            SQL = "SELECT * FROM qrySearchCustDomestic"
        Else
            SQL = "SELECT * FROM qrySearchCustDomestic WHERE [Name] like '" & Trim(Criteria) & "%' "
        End If
End If

If Me.Option2 Then
        If Len(Criteria) < 1 Then
            SQL = "SELECT * FROM qrySearchCustDomestic"
        Else
            SQL = "SELECT * FROM qrySearchCustDomestic WHERE [Ticket No] like '" & Trim(Criteria) & "%' "
        End If
End If


    'RsCust.CursorLocation = adUseClient
    Set RsCust = New ADODB.Recordset
    RsCust.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = RsCust

End Sub

Private Sub Option2_Click()
Call kulotHL(Me.Text1)
End Sub

Private Sub Text1_Change()
Call Disp_Cust(Me.Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo FailSafe_Error
If KeyCode = 13 Then
    If Me.Tag <> "over_ride" Then
            frmCashier.Text1 = Me.DataGrid1.Columns(0).Text
       Else
            frmSelectStatement.txtSearch = Me.DataGrid1.Columns(0).Text
    End If
End If
Exit Sub
FailSafe_Error:
    If Me.Tag <> "over_ride" Then
            frmCashier.Text1 = ""
    End If
End Sub
