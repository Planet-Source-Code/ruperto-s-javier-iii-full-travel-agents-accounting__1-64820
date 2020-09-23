VERSION 5.00
Begin VB.Form frmPO_INTL_insert 
   Caption         =   "Insert Pax Details"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlngIndex 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1155
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   480
      Left            =   6285
      TabIndex        =   5
      Top             =   1065
      Width           =   1740
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   480
      Left            =   4530
      TabIndex        =   4
      Top             =   1065
      Width           =   1740
   End
   Begin VB.TextBox txtTicketNo 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   585
      Width           =   3330
   End
   Begin VB.TextBox txtPaxName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   4440
   End
   Begin VB.Label Label2 
      Caption         =   "Ticket # :"
      Height          =   360
      Left            =   4695
      TabIndex        =   2
      Top             =   300
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Pax Name :"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   285
      Width           =   1890
   End
End
Attribute VB_Name = "frmPO_INTL_insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdInsert_Click()
Dim mylist As ListItem

If Me.Tag = "" Then
        With frmPO_INTL
             For i = 1 To .ListView1.ListItems.Count
                  If .ListView1.ListItems(i).SubItems(2) = Me.txtPaxName Then
                  MsgBox "This pax name was already inserted", vbInformation
                         Exit Sub
                  End If
             Next i
             Set mylist = .ListView1.ListItems.Add(, , "")
                   mylist.SubItems(1) = ""
                   mylist.SubItems(2) = Me.txtPaxName
                   mylist.SubItems(3) = Me.txtTicketNo
        End With
Exit Sub
End If

'==============================================================================================
If Me.Tag = "SA_INTL" Then
        With frmStatement_INTL
             For i = 1 To .ListView1.ListItems.Count
                  If .ListView1.ListItems(i).SubItems(2) = Me.txtPaxName Then
                  MsgBox "This pax name was already inserted", vbInformation
                         Exit Sub
                  End If
             Next i
             Set mylist = .ListView1.ListItems.Add(, , "")
                   mylist.SubItems(1) = ""
                   mylist.SubItems(2) = Me.txtPaxName
                   mylist.SubItems(3) = Me.txtTicketNo
        End With
Exit Sub
End If
'==============================================================================================
If Me.Tag = "SA_INTL_EDIT" Then
        With frmStatement_INTL
                .ListView1.ListItems(CLng(Me.txtlngIndex)).SubItems(2) = Me.txtPaxName
                .ListView1.ListItems(CLng(Me.txtlngIndex)).SubItems(3) = Me.txtTicketNo
        End With
End If
'==============================================================================================




End Sub

