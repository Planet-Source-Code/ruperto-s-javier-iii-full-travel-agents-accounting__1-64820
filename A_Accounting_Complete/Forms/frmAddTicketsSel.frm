VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddTicketsSel 
   Caption         =   "Select"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPODetailsID 
      Enabled         =   0   'False
      Height          =   405
      Left            =   6255
      TabIndex        =   9
      Top             =   555
      Width           =   1230
   End
   Begin VB.TextBox txtairline 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1050
      Width           =   3660
   End
   Begin VB.TextBox txtSelected 
      Enabled         =   0   'False
      Height          =   405
      Left            =   6255
      TabIndex        =   5
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   540
      Left            =   90
      TabIndex        =   4
      Top             =   4530
      Width           =   1395
   End
   Begin VB.TextBox txtPONumber 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3660
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   540
      Left            =   6105
      TabIndex        =   0
      Top             =   4545
      Width           =   1395
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   45
      TabIndex        =   3
      Top             =   1635
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POIDDetails"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "POID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Particulars"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "From"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "To"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Amount"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "PO Details ID"
      Height          =   330
      Left            =   4920
      TabIndex        =   10
      Top             =   570
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Airline / Shipping Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   150
      TabIndex        =   8
      Top             =   705
      Width           =   3525
   End
   Begin VB.Label Label2 
      Caption         =   "Row Selected"
      Height          =   330
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "PO #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   165
      TabIndex        =   1
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "frmAddTicketsSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPO As ADODB.Recordset

Private Sub cmdExit_Click()
Dim ask As Integer
ask = MsgBox("Are you finished?", vbYesNo + vbQuestion, "Confirm")
If ask = vbYes Then
    frmAddTickets.Tag = ""
       With frmAddTickets
            .txtStart = ""
            .txtEnd = ""
            .cboAirline.Enabled = False
            .cboTicketType.Enabled = False
            .cboAirline = ""
            .cboTicketType = ""
            .cmdAddSave.Caption = "Add"
       End With
    Unload Me
End If
End Sub

Private Sub cmdSelect_Click()
Dim lngIndex As Long
If Not CheckNull(Me.txtSelected) Then
lngIndex = CLng(Me.txtSelected)
  If Me.ListView1.ListItems(lngIndex).SubItems(7) = "Done" Then
        MsgBox "This ticket range was already Added", vbInformation
        Exit Sub
  Else
       With frmAddTickets
            .txtStart = Me.ListView1.ListItems(lngIndex).SubItems(3)
            .txtEnd = Me.ListView1.ListItems(lngIndex).SubItems(4)
            .cboAirline = Me.txtairline
            .cboTicketType = Me.ListView1.ListItems(lngIndex).SubItems(2)
           .cmdAddSave.Caption = "Save"
           frmAddTickets.Tag = "proceed"
       End With
  End If
End If
End Sub

Private Sub Form_Activate()
Call LoadValues(CDbl(Me.Tag))
End Sub

Private Sub Form_Load()
Call FormOnTop(Me.hWnd, True)


End Sub

Sub LoadValues(param)
Dim RsPODetails         As New ADODB.Recordset
Dim mySQL               As String
Dim ctr                 As Integer

    
Set RsPO = New ADODB.Recordset
SQL = "SELECT * FROM Tbl_PO_Domestic WHERE [PoID]=" & CLng(param)
With RsPO
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
       If .RecordCount < 1 Then
            .Close
           Set RsPO = Nothing
           Exit Sub
       End If
End With
    

With Me
        .txtPONumber = RsPO.Fields("Po Number").Value
        .txtairline = RsPO.Fields("Pay to").Value
        .Tag = RsPO.Fields("PoID").Value

'//=======================================================================
'//pull out data from details and load it to list view
'//=======================================================================
mySQL = "SELECT * FROM tbl_PODetails_Domestic WHERE [POid]=" & param
Me.ListView1.ListItems.Clear
         With RsPODetails
                .Open mySQL, cn, adOpenKeyset, adLockOptimistic
               If .RecordCount > 0 Then
                .MoveFirst
                    
                    ctr = 0
                    On Error Resume Next
                    Do While Not .EOF
                    ctr = ctr + 1
                        ListView1.ListItems.Add , , .Fields("POID_Details").Value
                        ListView1.ListItems.Item(ctr).SubItems(1) = .Fields("POID").Value
                        ListView1.ListItems.Item(ctr).SubItems(2) = .Fields("Particulars").Value
                        ListView1.ListItems.Item(ctr).SubItems(3) = .Fields("from").Value
                        ListView1.ListItems.Item(ctr).SubItems(4) = .Fields("to").Value
                        ListView1.ListItems.Item(ctr).SubItems(5) = .Fields("qty").Value
                        ListView1.ListItems.Item(ctr).SubItems(6) = Format(.Fields("Amount").Value, "###,##0.00")
                        ListView1.ListItems.Item(ctr).SubItems(7) = IIf(.Fields("encoded").Value = True, "Done", "Not Done")

                        .MoveNext
                    Loop
               End If
        End With

 
'//=======================================================================
End With

End Sub

Private Sub ListView1_Click()
If Me.ListView1.ListItems.Count > 0 Then
        Me.txtSelected = Me.ListView1.SelectedItem.Index
        Me.txtPODetailsID = ListView1.ListItems(CLng(txtSelected)).Text
End If
End Sub
