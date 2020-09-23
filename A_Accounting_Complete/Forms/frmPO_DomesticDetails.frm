VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmPO_DomesticDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PO Details"
   ClientHeight    =   3255
   ClientLeft      =   210
   ClientTop       =   675
   ClientWidth     =   10200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   4125
      Top             =   135
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
            Picture         =   "frmPO_DomesticDetails.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":0CDA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":1B2C
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":297E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":3258
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":3B32
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":440C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":4DD6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":56B0
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":59CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":62A4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":6B7E
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":7458
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":7772
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":804C
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":8926
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":9200
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":9ADA
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":A3B4
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":AC8E
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":B568
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":BE42
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":C71C
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":CFF6
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":D8D0
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":E1AA
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":EA84
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":F35E
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":FC38
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":10512
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":10DC8
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":116A2
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":11AF4
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":11F46
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":146F8
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO_DomesticDetails.frx":1597A
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   10140
      Begin VB.ComboBox cboAirline 
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   780
         Width           =   4680
      End
      Begin VB.ComboBox cboTicketType 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   780
         Width           =   4680
      End
      Begin VB.Label Label7 
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
         TabIndex        =   14
         Top             =   420
         Width           =   4200
      End
      Begin VB.Label Label6 
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
         TabIndex        =   13
         Top             =   435
         Width           =   4200
      End
   End
   Begin VB.TextBox txtQty 
      Height          =   330
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   1830
      Width           =   1845
   End
   Begin VB.TextBox txtTo 
      Height          =   330
      Left            =   3105
      TabIndex        =   5
      Text            =   "0"
      Top             =   1845
      Width           =   2115
   End
   Begin VB.TextBox txtFrom 
      Height          =   330
      Left            =   1065
      TabIndex        =   4
      Text            =   "0"
      Top             =   1830
      Width           =   1860
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7335
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   1815
      Width           =   2085
   End
   Begin LVbuttons.LaVolpeButton cmdExit 
      Height          =   480
      Left            =   5055
      TabIndex        =   2
      Top             =   2610
      Width           =   1620
      _ExtentX        =   2858
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
      MICON           =   "frmPO_DomesticDetails.frx":15C94
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
      Left            =   3375
      TabIndex        =   1
      Top             =   2610
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
      MICON           =   "frmPO_DomesticDetails.frx":15CB0
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   1455
      Width           =   765
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1485
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   1470
      Width           =   1785
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
      Left            =   7260
      TabIndex        =   3
      Top             =   1455
      Width           =   2145
   End
End
Attribute VB_Name = "frmPO_DomesticDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsInsert    As ADODB.Recordset
Dim myList      As ListItem
Dim SQL As String



Private Sub cboAirline_Click()
Dim Rs As New ADODB.Recordset
SQL = "SELECT  * FROM qryTickets WHERE [AirlineID]=" & FindAirlineID(Me.cboAirline)
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
Call FillTicketType
Rs.Close
End Sub
Function FindAirlineID(param) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & UCase(param) & "'"
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
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdInsert_Click()
                            
                            
If CheckNull(Me.cboAirline) Then
    MsgBox "Particulars should not be blank", vbCritical
    Me.cboAirline.SetFocus
    Exit Sub
End If

If CheckNull(Me.cboTicketType) Then
    MsgBox "Ticket should not be blank", vbCritical
    Me.cboTicketType.SetFocus
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

frmPO_Domestic.txtPayto = Me.cboAirline

Set myList = frmPO_Domestic.ListView1.ListItems.Add(, , "")
             myList.SubItems(1) = ""
             myList.SubItems(2) = Me.cboTicketType
             myList.SubItems(3) = Me.txtFrom
             myList.SubItems(4) = Me.txtTo
             myList.SubItems(5) = Me.txtQty
             myList.SubItems(6) = Me.txtAmount
 
With frmPO_Domestic
    .txtTotalAmount = Format(.SumListView, "###,##0.00")
End With

End Sub

Private Sub Form_Load()
Call FillAirline
Call FillTicketType
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

Private Sub txtFrom_Change()
If Not IsNumeric(Me.txtFrom) Then
    Me.txtFrom = "0"
    Call kulotHL(Me.txtFrom)
    Exit Sub
End If
Me.txtQty = CDbl(Me.txtTo) - CDbl(Me.txtFrom)
End Sub

Private Sub txtTo_Change()
If Not IsNumeric(Me.txtTo) Then
    Me.txtTo = "0"
    Call kulotHL(Me.txtTo)
    Exit Sub
End If
Me.txtQty = (CDbl(Me.txtTo) - CDbl(Me.txtFrom)) + 1
End Sub



Sub FillAirline()
Dim Rst As New ADODB.Recordset
SQL = "SELECT  * FROM tbl_Airline WHERE [AirlineName]<>'NONE' ORDER by [AirlineName] ASC "
        Me.cboAirline.Clear
        
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboAirline.AddItem .Fields(1).Value
                .MoveNext
            Loop
         Me.cboAirline.ListIndex = 0
        End If
       .Close
      Set Rst = Nothing
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

Set Rst = Nothing

End With

End Sub

