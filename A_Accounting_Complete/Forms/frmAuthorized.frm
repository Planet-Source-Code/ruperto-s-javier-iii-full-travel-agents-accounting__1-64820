VERSION 5.00
Begin VB.Form frmAuthorized 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Authorized Representative"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      Height          =   480
      Left            =   285
      TabIndex        =   0
      Top             =   525
      Width           =   5730
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   450
      Left            =   3960
      TabIndex        =   1
      Top             =   1140
      Width           =   2040
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Print"
      Height          =   450
      Left            =   1935
      TabIndex        =   2
      Top             =   1140
      Width           =   2040
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name of authorized Representative"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAuthorized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim Rst As New ADODB.Recordset
Dim Rpt As New RptPO
SQL = "SELECT * FROM Tbl_PO_Domestic WHERE [Po Number]='" & frmPO_Domestic.txtPONumber & "'"
With Rst
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
            SQL = "SELECT * FROM tbl_PODetails_Domestic WHERE [POID]=" & .Fields("POID").Value
                With Rpt
                        .DataControl1.Connection = cn
                        .DataControl1.Source = SQL
                        .lblAuthorized.Caption = UCase(Me.Text1)
                        .lblCompany.Caption = "TO :" & frmPO_Domestic.txtPayto
                        .lblSpecimen = UCase(Me.Text1)
                        .Show 1
                End With

    Else
                MsgBox "This PO# was not save.", vbApplicationModal, "Error Printing"
    End If
    .Close
    Set Rst = Nothing
End With

End Sub
