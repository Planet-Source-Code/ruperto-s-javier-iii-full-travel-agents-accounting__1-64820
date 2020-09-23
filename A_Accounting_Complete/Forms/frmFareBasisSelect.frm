VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFareBasisSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fare Basis Select"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   390
      Left            =   135
      TabIndex        =   2
      Top             =   3825
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   4620
      TabIndex        =   1
      Top             =   3810
      Width           =   1350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3570
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   6297
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "FareBasisID"
         Caption         =   "FareBasisID"
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
         DataField       =   "FareBasis"
         Caption         =   "FareBasis"
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
         DataField       =   "FareBasisAmount"
         Caption         =   "FareBasis Amount"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3165.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1980.284
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFareBasisSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsFareBasis As ADODB.Recordset
Dim RsRouteFareBasis As ADODB.Recordset
Dim RsTempFareBasis As ADODB.Recordset
Dim SQL As String


Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdInsert_Click()
On Error GoTo FailSafe_Error
Dim ask As Integer
SQL = "SELECT * FROM tbl_TmpFareBasis"

If CheckDupFB(Me.DataGrid1.Columns(0).Text) Then
    MsgBox "The Farebasis you selected already exist on this route", vbCritical
    Exit Sub
End If

ask = MsgBox("Insert selected Farebasis?", vbQuestion + vbYesNo, "ELS")
If ask = vbYes Then
    Set RsTempFareBasis = New ADODB.Recordset
            With RsTempFareBasis
                        .Open SQL, cn, adOpenKeyset, adLockOptimistic
                        .AddNew
                            .Fields(0).Value = Me.DataGrid1.Columns(0).Text
                            .Fields(1).Value = Me.DataGrid1.Columns(1).Text
                            .Fields(2).Value = Me.DataGrid1.Columns(2).Text
                            .Fields(3).Value = CLng(frmSetTicketPricing.DataGrid2.Columns(0).Text)
                        .Update
                        .Close
                        MsgBox "Fare Basis inserted...", vbInformation, "ELS"
                     Set RsTempFareBasis = Nothing
            End With
End If
frmSetTicketPricing.DisplayFB
Exit Sub
FailSafe_Error:
MsgBox "There was an error in inserting fare basis", vbInformation, "ELS"
End Sub

Private Sub Form_Load()
SQL = "SELECT * FROM tbl_FareBasis ORDER by [FareBasis] ASC"
Set RsFareBasis = New ADODB.Recordset
RsFareBasis.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = RsFareBasis

SQL = "SELECT * FROM tbl_RoutePricing"
Set RsRouteFareBasis = New ADODB.Recordset
RsRouteFareBasis.Open SQL, cn, adOpenKeyset, adLockOptimistic

End Sub

Function CheckDupFB(Param) As Boolean
Dim RstTemp As New ADODB.Recordset
SQL = "SELECT * from tbl_TmpFareBasis WHERE [FareBasisID]=" & CLng(Param)
With RstTemp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                CheckDupFB = True
                Else
                CheckDupFB = False
            End If
            .Close
        Set RstTemp = Nothing
End With
End Function
