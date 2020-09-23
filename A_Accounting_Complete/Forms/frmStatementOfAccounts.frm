VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStatementOfAccounts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Statement"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   7785
      TabIndex        =   11
      Text            =   "Combo5"
      Top             =   2145
      Width           =   4650
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Text            =   "Combo4"
      Top             =   2130
      Width           =   4815
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New..."
      Height          =   285
      Left            =   6810
      TabIndex        =   7
      Top             =   1260
      Width           =   1185
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   1245
      Width           =   4815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   4815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2250
      Left            =   225
      TabIndex        =   2
      Top             =   2745
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   3969
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Available Routes"
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "RoutePricingID"
         Caption         =   "RoutePricingID"
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
         DataField       =   "RouteID"
         Caption         =   "RouteID"
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
         DataField       =   "From"
         Caption         =   "From"
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
         DataField       =   "To"
         Caption         =   "To"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "Gross Fare"
         Caption         =   "Gross Fare"
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
      BeginProperty Column09 
         DataField       =   "Insurance"
         Caption         =   "Insurance"
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
      BeginProperty Column10 
         DataField       =   "Commision"
         Caption         =   "Commision"
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
      BeginProperty Column11 
         DataField       =   "ASF"
         Caption         =   "ASF"
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
      BeginProperty Column12 
         DataField       =   "Net Fare"
         Caption         =   "Net Fare"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2459.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   4815
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   3030
      Left            =   225
      TabIndex        =   8
      Top             =   5235
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   5345
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "To :"
      Height          =   315
      Left            =   7230
      TabIndex        =   12
      Top             =   2145
      Width           =   690
   End
   Begin VB.Label Label4 
      Caption         =   "From :"
      Height          =   315
      Left            =   1065
      TabIndex        =   10
      Top             =   2130
      Width           =   690
   End
   Begin VB.Label Label3 
      Caption         =   "Agent Name :"
      Height          =   315
      Left            =   195
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Type of Ticket :"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   735
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Shipping / Airline Name:"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmStatementOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim SQL As String

Sub FillComboShip()
Dim tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline"
With tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo1.Clear
                Do While Not .EOF
                    Me.Combo1.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With



SQL = "SELECT * FROM Query1"
Set tmp = New ADODB.Recordset
With tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo4.Clear
                Do While Not .EOF
                    Me.Combo4.AddItem .Fields(0).Value
                .MoveNext
                Loop
           End If
End With



SQL = "SELECT * FROM Query2"
Set tmp = New ADODB.Recordset
With tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo5.Clear
                Do While Not .EOF
                    Me.Combo5.AddItem .Fields(0).Value
                .MoveNext
                Loop
           End If
End With
End Sub

Sub FillComboTicketType()
Dim tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_TicketType"
With tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo2.Clear
                Do While Not .EOF
                    Me.Combo2.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With
End Sub

Sub displayGrid()
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            Set Me.DataGrid1.DataSource = Rs
End With
End Sub

Private Sub Combo1_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [AirlineName]='" & Me.Combo1 & "' ORDER by [AirlineName]"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        
          
            Set Me.DataGrid1.DataSource = Rs
           
End With


Set Rs1 = New ADODB.Recordset
SQL = "SELECT * from qryCboTickets WHERE [AirlineName]='" & Me.Combo1 & "'"
With Rs1
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        
          If .RecordCount > 0 Then
                Me.Combo2.Clear
                Do While Not .EOF
                    Me.Combo2.AddItem .Fields(3).Value
                .MoveNext
                Loop
                Else
                Me.Combo2.Clear
          End If
            
           
End With

End Sub

Private Sub Combo2_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [Ticket Type]='" & Me.Combo2 & "' order by from"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        
          
            Set Me.DataGrid1.DataSource = Rs
           
End With

End Sub

Private Sub Combo4_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [FROM]='" & Me.Combo4 & "' AND [TO]='" & Me.Combo5 & "' order by from"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        
          
            Set Me.DataGrid1.DataSource = Rs
           
End With
End Sub

Private Sub Combo5_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * from qryRoutePricing WHERE [FROM]='" & Me.Combo4 & "' AND [TO]='" & Me.Combo5 & "' order by from"
With Rs
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        
          
            Set Me.DataGrid1.DataSource = Rs
           
End With
End Sub

Private Sub Form_Load()

Call displayGrid
Call FillComboShip
Call FillComboTicketType
End Sub
