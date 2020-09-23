VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFareBasis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fare Basis"
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
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   270
      Left            =   1845
      TabIndex        =   9
      Top             =   75
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   390
      Left            =   1485
      TabIndex        =   8
      Top             =   3825
      Width           =   1350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   390
      Left            =   135
      TabIndex        =   7
      Top             =   3825
      Width           =   1350
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   390
      Left            =   2820
      TabIndex        =   6
      Top             =   3825
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   4590
      TabIndex        =   5
      Top             =   3825
      Width           =   1350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2520
      Left            =   150
      TabIndex        =   4
      Top             =   1215
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   4445
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3795.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1980.284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2625
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   705
      Width           =   3300
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4515
      MaxLength       =   3
      TabIndex        =   0
      Top             =   225
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   3
      Top             =   750
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Fare Basis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   1
      Top             =   270
      Width           =   1425
   End
End
Attribute VB_Name = "frmFareBasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsFareBasis As ADODB.Recordset
Dim SQL As String

Private Sub cmdDelete_Click()
On Error GoTo FailSafe
Dim ask As Integer
SQL = "DELETE * FROM tbl_FareBasis WHERE [FareBasisID]=" & Me.DataGrid1.Columns(0).Text
ask = MsgBox("Sure you want to delete this?", vbCritical + vbYesNo, "ELS")
If ask = vbYes Then
    cn.Execute SQL
    RsFareBasis.Requery
    MsgBox "One record deleted", vbInformation
End If
Exit Sub
FailSafe:
MsgBox "There was an error in deleting fare basis..."
End Sub

Private Sub cmdEdit_Click()
Me.Check1.Value = 1
Me.Text1 = Me.DataGrid1.Columns(1).Value
Me.Text2 = Me.DataGrid1.Columns(2).Value
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If CheckNull(Me.Text1) Then
        MsgBox "Fare Basis should not be blank", vbCritical, "ELS"
        With Me.Text1
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
        End With
        Exit Sub
End If



If CDbl(Me.Text2) = 0 Then
        MsgBox "Amount should be greater zero", vbCritical, "ELS"
        With Me.Text2
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
        End With
        Exit Sub
End If

If Not IsNumeric(Me.Text2) Then
        MsgBox "Amount requires numeric value", vbCritical, "ELS"
        With Me.Text2
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
        End With
        Exit Sub
End If

        With RsFareBasis


            If Me.Check1.Value Then
                    .Find "[FareBasisID]=" & Me.DataGrid1.Columns(0).Text
            Else
                If CheckDupFB(Me.Text1) Then
                    MsgBox "This Fare Basis already exist pls create new one", vbCritical, "ELS"
                    Exit Sub
                End If

                    .AddNew
            End If
            If Me.Check1.Value = 1 Then
                    MsgBox "New price rate updated...", vbInformation, "ELS"
            End If
                    .Fields(1).Value = UCase(Me.Text1)
                    .Fields(2).Value = CDbl(Me.Text2)
                    .Update
        End With
                    Me.Check1.Value = 0


End Sub

Private Sub Form_Load()
SQL = "SELECT * FROM tbl_FareBasis ORDER by [FareBasis] ASC"
Set RsFareBasis = New ADODB.Recordset
RsFareBasis.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = RsFareBasis
End Sub

Function CheckDupFB(Param) As Boolean
Dim RstTemp As New ADODB.Recordset
SQL = "SELECT * from tbl_FareBasis WHERE [FareBasis]='" & Param & "'"
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
