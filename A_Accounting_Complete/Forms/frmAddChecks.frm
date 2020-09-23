VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmAddChecks 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Cheques"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10695
      TabIndex        =   15
      Top             =   -15
      Width           =   10755
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmAddChecks.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Adding new data of checks"
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
         TabIndex        =   17
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Checks"
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
         TabIndex        =   16
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Height          =   1020
      Left            =   75
      ScaleHeight     =   960
      ScaleWidth      =   10575
      TabIndex        =   10
      Top             =   2340
      Width           =   10635
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   345
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Cheque No. :"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "End Cheque No. :"
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
         TabIndex        =   13
         Top             =   390
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Which Bank and Account Number"
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
      Left            =   75
      TabIndex        =   5
      Top             =   1035
      Width           =   10635
      Begin VB.ComboBox cboAccNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   780
         Width           =   4680
      End
      Begin VB.ComboBox cboBank 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   780
         Width           =   4680
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         TabIndex        =   8
         Top             =   435
         Width           =   4200
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         TabIndex        =   7
         Top             =   405
         Width           =   4200
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   9765
      Top             =   5145
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
            Picture         =   "frmAddChecks.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":15A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":23F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":3248
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":3B22
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":43FC
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":4CD6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":56A0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":5F7A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":6294
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":6B6E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":7448
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":7D22
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":803C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":8916
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":91F0
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":9ACA
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":A3A4
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":AC7E
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":B558
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":BE32
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":C70C
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":CFE6
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":D8C0
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":E19A
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":EA74
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":F34E
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":FC28
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":10502
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":10DDC
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":11692
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":11F6C
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":123BE
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":12810
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddChecks.frx":14FC2
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   90
      TabIndex        =   1
      Top             =   6990
      Width           =   10590
      Begin LVbuttons.LaVolpeButton cmdAddSave 
         Height          =   480
         Left            =   105
         TabIndex        =   2
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
         MICON           =   "frmAddChecks.frx":16244
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
         Left            =   8895
         TabIndex        =   3
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
         MICON           =   "frmAddChecks.frx":16260
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
         TabIndex        =   4
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
         MICON           =   "frmAddChecks.frx":1627C
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
      Height          =   3555
      Left            =   90
      TabIndex        =   0
      Top             =   3420
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6271
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "CheckID"
         Caption         =   "CheckID"
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
         DataField       =   "CheckNo"
         Caption         =   "CheckNo"
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
         DataField       =   "BankID"
         Caption         =   "BankID"
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
         DataField       =   "Bank Name"
         Caption         =   "Bank Name"
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
         DataField       =   "AccNo"
         Caption         =   "AccNo"
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
         AllowFocus      =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2459.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim RsChecks As New ADODB.Recordset
Dim SQL As String
Dim OldBookMark



Private Sub cboBank_Click()
    
    FillAccNo
End Sub

Private Sub cmdAddSave_Click()
If Me.cmdAddSave.Caption = "Add" Then
    Me.cmdAddSave.Caption = "Save"
    
    Me.cboBank.Enabled = True
    Me.cboAccNo.Enabled = True
    Me.txtStart.Enabled = True
    Me.txtEnd.Enabled = True
    
Else
    If CheckNull(Me.cboBank) Then
        MsgBox "The Bank Name Should Not be Blank"
        Exit Sub
    End If
    
    If CheckNull(Me.cboAccNo) Then
        MsgBox "The Bank account Should Not be Blank"
        Exit Sub
    End If
    
    If CheckNull(Me.txtStart) Then
        MsgBox "The Starting Ticket No. Should Not be Blank"
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
    
    
    Dim i As Double
    For i = CDbl(Me.txtStart) To CDbl(Me.txtEnd) Step 1
        If Not DupCheck(i) Then
                With RsChecks
                cn.BeginTrans
                        .AddNew
                        .Fields(1).Value = CDbl(i)
                        .Fields(2).Value = FindBankID(Me.cboBank)
                        .Fields(3).Value = Me.cboAccNo
                        .Fields(4).Value = "Un-Used"
                        .Fields(5).Value = Format(Now, "mm/dd/yyyy")
                        
                        .Update
                cn.CommitTrans
                End With
         End If
    Next i
    Me.cmdAddSave.Caption = "Add"
    
Me.cboBank.Enabled = False
Me.cboAccNo.Enabled = False
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
    Rs.Requery
    
    
End If

Exit Sub

ErrExit:
cn.RollbackTrans

End Sub

Private Sub cmdAddSave_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim ask As Integer

ask = MsgBox("Are you sure you want to remove this Check No?", vbCritical + vbYesNo)
If ask = vbYes Then
SQL = "DELETE * FROM tbl_checks WHERE [CheckID]=" & Me.DataGrid1.Columns(0).Text
cn.Execute SQL
OldBookMark = Rs.Bookmark
Rs.Requery
MsgBox "One check(s) deleted"
Rs.Bookmark = OldBookMark
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set Rs = New ADODB.Recordset
SQL = "SELECT  * FROM qryChecks"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Set Me.DataGrid1.DataSource = Rs

Set RsChecks = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_checks"
RsChecks.Open SQL, cn, adOpenKeyset, adLockOptimistic


Call FillBank
Call FillAccNo
End Sub

Function DupCheck(check) As Boolean
Dim tmp As ADODB.Recordset

Set tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_checks WHERE [CheckNo]='" & UCase(check) & "' AND [BankID]=" & FindBankID(Me.cboBank)
With tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            DupCheck = True
        Else
            DupCheck = False
        End If
      .Close
    Set tmp = Nothing
End With

End Function


Sub FillBank()
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks"
        Me.cboBank.Clear
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboBank.AddItem .Fields(1).Value
                .MoveNext
            Loop
        End If
       .Close
     Set Rst = Nothing
End With
End Sub


Sub FillAccNo()
Dim Rst As New ADODB.Recordset
SQL = "SELECT DISTINCT [Account Number] FROM tbl_AccountsSetting WHERE [BankID]=" & FindBankID(Me.cboBank)
       Me.cboAccNo.Clear
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboAccNo.AddItem .Fields("Account Number").Value
                .MoveNext
            Loop
        End If
        .Close
     Set Rst = Nothing

End With

End Sub

Function FindBankID(Param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks WHERE [Bank Name]='" & UCase(Param) & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindBankID = .Fields(0).Value
          Else
              FindBankID = -1
        End If
     .Close
   Set Rst = Nothing
End With
End Function

