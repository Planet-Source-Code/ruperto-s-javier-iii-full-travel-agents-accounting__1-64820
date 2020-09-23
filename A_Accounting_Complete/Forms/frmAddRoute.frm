VERSION 5.00
Begin VB.Form frmAddRoute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Route"
   ClientHeight    =   1545
   ClientLeft      =   330
   ClientTop       =   990
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   450
      Left            =   4395
      TabIndex        =   4
      Top             =   1005
      Width           =   1590
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Save"
      Height          =   450
      Left            =   75
      TabIndex        =   3
      Top             =   1005
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Route"
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   5940
      Begin VB.TextBox txtRoute 
         Height          =   375
         Left            =   1395
         TabIndex        =   2
         Top             =   330
         Width           =   4365
      End
      Begin VB.Label Label1 
         Caption         =   "Route :"
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   375
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAddRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim SQL As String
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSet_Click()

If Len(Me.txtRoute) < 1 Then
    MsgBox "Route should not be blank!", vbCritical
    With Me.txtRoute
        .SelLength = 0
        .SelStart = Len(.Text)
        .SetFocus
        Exit Sub
    End With

End If

With Rs
  If Me.Tag = "new" Then
    .AddNew
    Else
    .Find "[AvailableRouteID]=" & frmTicketType.DataGrid2.Columns(0).Text
  End If
    .Fields(1).Value = frmTicketType.DataGrid1.Columns(0).Text
    .Fields(2).Value = UCase(Me.txtTicketType)
    
    .Update
End With
frmTicketType.DisplayTicket
End Sub

Private Sub Form_Activate()
On Error GoTo ErrExit
Me.Text1 = frmTicketType.DataGrid1.Columns(1).Text
Set Rs = New ADODB.Recordset
If LCase(Me.Tag) = "edit" Then
        SQL = "SELECT * FROM tbl_TicketType WHERE [TicketTypeID]=" & frmTicketType.DataGrid2.Columns(0).Text
    Else
        SQL = "SELECT * FROM tbl_TicketType"
End If
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

If LCase(Me.Tag) = "edit" Then
    Me.txtTicketType = Rs.Fields(2).Value
   
End If
Exit Sub
ErrExit:
Me.Tag = "new"
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_TicketType"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub txtTicketType_GotFocus()
With txtTicketType
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
End With
End Sub
