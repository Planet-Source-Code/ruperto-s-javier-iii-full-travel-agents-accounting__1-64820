VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmAddPassengerType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Passenger Type"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   5940
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":0CDA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":1B2C
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":297E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":3258
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":3B32
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":440C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":4DD6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":56B0
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":59CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":62A4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":6B7E
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":7458
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":7772
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":804C
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":8926
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":9200
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":9ADA
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":A3B4
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":AC8E
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":B568
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":BE42
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":C71C
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":CFF6
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":D8D0
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":E1AA
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":EA84
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":F35E
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":FC38
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":10512
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":10DC8
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":116A2
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":11AF4
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":11F46
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddPassengerType.frx":146F8
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7275
      TabIndex        =   9
      Top             =   0
      Width           =   7335
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmAddPassengerType.frx":1597A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Adding new data of passenger"
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
         TabIndex        =   11
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Passenger Type"
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
         TabIndex        =   10
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.TextBox txtDiscount 
      Height          =   300
      Left            =   2070
      TabIndex        =   7
      Top             =   1425
      Width           =   2850
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   45
      TabIndex        =   3
      Top             =   6600
      Width           =   7185
      Begin LVbuttons.LaVolpeButton cmdAddSave 
         Height          =   480
         Left            =   105
         TabIndex        =   4
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Add"
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
         MICON           =   "frmAddPassengerType.frx":16244
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "13"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdExit 
         Height          =   480
         Left            =   5505
         TabIndex        =   5
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
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
         MICON           =   "frmAddPassengerType.frx":16260
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
      Begin LVbuttons.LaVolpeButton cmdDelete 
         Height          =   480
         Left            =   1605
         TabIndex        =   6
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "Delete"
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
         MICON           =   "frmAddPassengerType.frx":1627C
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "14"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   45
      TabIndex        =   2
      Top             =   1875
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   8281
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "PassengerTypeID"
         Caption         =   "PassengerTypeID"
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
         DataField       =   "PassengerType"
         Caption         =   "PassengerType"
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
         DataField       =   "Discount Percent"
         Caption         =   "Discount Percent (%)"
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
         AllowFocus      =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3555.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1769.953
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPassengerType 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2070
      TabIndex        =   0
      Top             =   1080
      Width           =   2865
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount in Percent:"
      Height          =   285
      Left            =   105
      TabIndex        =   8
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger Type"
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   1095
      Width           =   1890
   End
End
Attribute VB_Name = "frmAddPassengerType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim SQL As String
Dim OldBookMark

Private Sub cmdAddSave_Click()
If Me.cmdAddSave.Caption = "Add" Then
    Me.cmdAddSave.Caption = "Save"
    txtPassengerType.Enabled = True
    
    With Me.txtPassengerType
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
Else
    If CheckNull(Me.txtPassengerType) Then
        MsgBox "The passenger Type Should Not be Blank"
        Exit Sub
    End If
    
    If CheckNull(Me.txtDiscount) Then
        MsgBox "Discount Should Not be Blank"
        Exit Sub
    End If
    
    If DupShipAirline Then
        MsgBox "This Passenger Type already in the database.."
            With Me.txtPassengerType
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
        Exit Sub
        
    End If
    With Rs
            .AddNew
            .Fields(1).Value = UCase(Me.txtPassengerType)
            .Fields(2).Value = CDbl(Me.txtDiscount)
            .Update
    End With
    Me.cmdAddSave.Caption = "Add"
    txtPassengerType.Enabled = False
End If


End Sub

Private Sub cmdAddSave_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim ask As Integer

ask = MsgBox("Are you sure you want to remove this Passenger Type?", vbCritical + vbYesNo)
If ask = vbYes Then
SQL = "DELETE * FROM tbl_PassengerType WHERE [PassengerTypeID]=" & Me.DataGrid1.Columns(0).Text
cn.Execute SQL
OldBookMark = Rs.Bookmark
Rs.Requery
MsgBox "One Passenger Type deleted"
Rs.Bookmark = OldBookMark
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_PassengerType"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

Set Me.DataGrid1.DataSource = Rs
End Sub


Function DupPassengerType() As Boolean
Dim Tmp As ADODB.Recordset

Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_PassengerType WHERE [PassengerTypeID]='" & UCase(Me.txtPassengerType) & "'"
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            DupPassengerType = True
        Else
            DupPassengerType = False
        End If
End With

End Function


Private Sub txtAirline_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cmdAddSave.SetFocus
End If
End Sub
