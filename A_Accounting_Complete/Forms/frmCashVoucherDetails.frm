VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmCashVoucherDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Details"
   ClientHeight    =   1710
   ClientLeft      =   210
   ClientTop       =   675
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   435
      Left            =   5025
      TabIndex        =   6
      Top             =   495
      Width           =   1260
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   4620
      Top             =   1170
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
            Picture         =   "frmCashVoucherDetails.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":0CDA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":1B2C
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":297E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":3258
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":3B32
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":440C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":4DD6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":56B0
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":59CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":62A4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":6B7E
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":7458
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":7772
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":804C
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":8926
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":9200
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":9ADA
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":A3B4
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":AC8E
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":B568
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":BE42
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":C71C
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":CFF6
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":D8D0
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":E1AA
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":EA84
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":F35E
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":FC38
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":10512
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":10DC8
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":116A2
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":11AF4
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":11F46
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":146F8
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashVoucherDetails.frx":1597A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6315
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   495
      Width           =   2085
   End
   Begin VB.TextBox txtParticular 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   4935
   End
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   7020
      TabIndex        =   3
      Top             =   1170
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
      MICON           =   "frmCashVoucherDetails.frx":15C94
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
   Begin LVbuttons.LaVolpeButton cmdInsert 
      Height          =   480
      Left            =   5325
      TabIndex        =   2
      Top             =   1170
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Insert"
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCashVoucherDetails.frx":15CB0
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "22"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6330
      TabIndex        =   5
      Top             =   105
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   135
      Width           =   2145
   End
End
Attribute VB_Name = "frmCashVoucherDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsInsert    As ADODB.Recordset
Dim mylist      As ListItem
Dim SQL As String



Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
If frmCashVoucher.OptDomestic Then
    
Else
frmPO_INTL_Find.Tag = "insert po"
frmPO_INTL_Find.Show 1
End If
End Sub

Private Sub cmdInsert_Click()


                            
                            
If CheckNull(Me.txtParticular) Then
    MsgBox "Particulars should not be blank", vbCritical
    Me.txtParticular.SetFocus
    Exit Sub
End If

If Not IsNumeric((Me.txtAmount)) Then
    MsgBox "Invalid Amount", vbCritical
    With Me.txtAmount
            .SelLength = 0
            .SelStart = Len(.Text)
            .SetFocus
    End With
    Exit Sub
End If

If CDbl(Me.txtAmount) <= 0 Then
    MsgBox "Invalid Amount", vbCritical
    With Me.txtAmount
            .SelLength = 0
            .SelStart = Len(.Text)
            .SetFocus
    End With
    Exit Sub
End If



Set mylist = frmCashVoucher.ListView1.ListItems.Add(, , "")
             mylist.SubItems(1) = Me.txtParticular
             mylist.SubItems(2) = Format(Me.txtAmount, "###,##0.00")
 
 
'Set RsInsert = New ADODB.Recordset
'SQL = "SELECT * FROM tbl_VoucherDetails"
'With RsInsert
'            .Open SQL, cn, adOpenKeyset, adLockOptimistic
'            .AddNew
 '           .Fields(1).Value = Me.Tag
 '           .Fields(2).Value = UCase(Me.txtParticular)
 '           .Fields(3).Value = CDbl(Me.txtAmount)
'            .Update
'End With
With frmCashVoucher
    '.RefreshGrid (Me.Tag)
    '.UpDateVoucher (Me.Tag)
    .txtTotalAmount = Format(.SumListView, "###,##0.00")
End With
With Me.txtParticular
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
End With
End Sub

Private Sub txtAmount_GotFocus()
With Me.txtAmount
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
End With
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtAmount_LostFocus()
txtAmount = Format(Me.txtAmount, "###,##0.00")
End Sub

Private Sub txtParticular_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter KeyCode
End Sub

Private Sub txtParticular_LostFocus()
txtParticular = UCase(txtParticular)
End Sub
