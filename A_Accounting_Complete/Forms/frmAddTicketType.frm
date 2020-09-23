VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmAddTicketType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAddTicketType.frx":0000
      Left            =   1200
      List            =   "frmAddTicketType.frx":000A
      TabIndex        =   11
      Top             =   1740
      Width           =   4215
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   4980
      Top             =   6015
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
            Picture         =   "frmAddTicketType.frx":0026
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":0D00
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":1B52
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":29A4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":327E
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":3B58
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":4432
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":4DFC
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":56D6
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":59F0
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":62CA
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":6BA4
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":747E
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":7798
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":8072
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":894C
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":9226
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":9B00
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":A3DA
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":ACB4
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":B58E
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":BE68
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":C742
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":D01C
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":D8F6
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":E1D0
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":EAAA
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":F384
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":FC5E
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":10538
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":10DEE
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":116C8
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":11B1A
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":11F6C
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTicketType.frx":1471E
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   6255
      TabIndex        =   7
      Top             =   0
      Width           =   6315
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Ticket Type"
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
         TabIndex        =   9
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Adding new data of Ticket Type"
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
         TabIndex        =   8
         Top             =   600
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmAddTicketType.frx":159A0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   15
      TabIndex        =   3
      Top             =   7740
      Width           =   6240
      Begin LVbuttons.LaVolpeButton cmdAddSave 
         Height          =   480
         Left            =   120
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
         MICON           =   "frmAddTicketType.frx":1626A
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
         Left            =   4620
         TabIndex        =   5
         Top             =   210
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
         MICON           =   "frmAddTicketType.frx":16286
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
         Left            =   1620
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
         MICON           =   "frmAddTicketType.frx":162A2
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
      Height          =   5310
      Left            =   15
      TabIndex        =   2
      Top             =   2400
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   9366
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "AirlineShippingLine"
         Caption         =   "For Ship / Airline"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2805.166
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTicketType 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   1140
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Type"
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   1800
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Type"
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   1125
      Width           =   1890
   End
End
Attribute VB_Name = "frmAddTicketType"
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
    Me.txtTicketType.Enabled = True
    
    With Me.txtTicketType
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
    End With
Else
    If CheckNull(Me.txtTicketType) Then
        MsgBox "The Ticket type Should Not be Blank"
        Me.txtTicketType.SetFocus
        Exit Sub
    End If
    
    If DupTicket Then
        MsgBox "This Ticket type already in the database.."
            With Me.txtTicketType
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
            End With
        Exit Sub
        
    End If
    With Rs
            .AddNew
            .Fields(1).Value = UCase(Me.txtTicketType)
            .Fields(2).Value = UCase(Me.Combo1.Text)
            .Update
    End With
    Me.cmdAddSave.Caption = "Add"
    Me.txtTicketType.Enabled = False
End If


End Sub

Private Sub cmdAddSave_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim ask As Integer

ask = MsgBox("Are you sure you want to remove this Ticket Type?", vbCritical + vbYesNo)
If ask = vbYes Then
SQL = "DELETE * FROM tbl_TicketType WHERE [TicketTypeID]=" & Me.DataGrid1.Columns(0).Text
cn.Execute SQL
OldBookMark = Rs.Bookmark
Rs.Requery
MsgBox "One Ticket Type deleted..."
Rs.Bookmark = OldBookMark
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Me.txtTicketType = Me.DataGrid1.Columns(1).Text
'Me.Combo1 = Me.DataGrid1.Columns(2).Text
End Sub

Private Sub Form_Load()
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_TicketType"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

Set Me.DataGrid1.DataSource = Rs

Me.Combo1.ListIndex = 0
Me.txtTicketType = Me.DataGrid1.Columns(1).Text
Me.Combo1 = Me.DataGrid1.Columns(2).Text
End Sub


Function DupTicket() As Boolean
Dim Tmp As ADODB.Recordset

Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_TicketType WHERE [Ticket Type]='" & UCase(Me.txtTicketType) & "'"
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            DupTicket = True
        Else
            DupTicket = False
        End If
End With

End Function


Private Sub txtTicketType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cmdAddSave.SetFocus
End If
End Sub
