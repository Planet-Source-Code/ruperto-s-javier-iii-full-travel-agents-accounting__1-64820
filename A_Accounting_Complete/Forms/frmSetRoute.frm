VERSION 5.00
Begin VB.Form frmSetRoute 
   Caption         =   "Route"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Add"
      Height          =   420
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   1410
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   2190
      TabIndex        =   17
      Top             =   3870
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Height          =   3180
      Left            =   75
      TabIndex        =   2
      Top             =   645
      Width           =   3555
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1215
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   2580
         Width           =   2130
      End
      Begin VB.TextBox txtfare 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   2205
         Width           =   2130
      End
      Begin VB.TextBox txtvat 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1215
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1800
         Width           =   2130
      End
      Begin VB.TextBox txtcomm 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1215
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1395
         Width           =   2130
      End
      Begin VB.TextBox txtinsurance 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1215
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1020
         Width           =   2130
      End
      Begin VB.TextBox txtasf 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1215
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox txtfee 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1215
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   195
         Width           =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "Total :"
         Height          =   345
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Fare :"
         Height          =   345
         Index           =   6
         Left            =   165
         TabIndex        =   13
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Vat :"
         Height          =   345
         Index           =   5
         Left            =   165
         TabIndex        =   11
         Top             =   1935
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Commision :"
         Height          =   345
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Insurance :"
         Height          =   345
         Index           =   3
         Left            =   135
         TabIndex        =   7
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "ASF :"
         Height          =   345
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   705
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Terminal Fee :"
         Height          =   345
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtRoute 
      Height          =   405
      Left            =   720
      TabIndex        =   1
      Top             =   180
      Width           =   2700
   End
   Begin VB.Label Label1 
      Caption         =   "Route :"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   255
      Width           =   600
   End
End
Attribute VB_Name = "frmSetRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim SQL As String


Private Sub cmdAddSave_Click()
    
    
        With Rs
        If Me.cmdAddSave.Caption = "Add" Then
                    .AddNew
                    Me.cmdAddSave.Caption = "Save"
        Else
                    Me.cmdAddSave.Caption = "Add"
        End If
                    .Fields(1).Value = CDbl(frmTicketType.DataGrid2.Columns(0).Text)
                    .Fields(2).Value = Me.TxtRoute
                    .Fields(3).Value = CDbl(Me.txtfee)
                    .Fields(4).Value = CDbl(Me.txtasf)
                    .Fields(5).Value = CDbl(Me.txtinsurance)
                    .Fields(6).Value = CDbl(Me.txtcomm)
                    .Fields(7).Value = CDbl(Me.txtvat)
                    .Fields(8).Value = CDbl(Me.txtfare)
                    .Fields(9).Value = CDbl(Me.txttotal)
                    .Update
        End With

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Route"
With Rs
.Open SQL, cn, adOpenKeyset, adLockOptimistic
End With
End Sub

Private Sub txtasf_Change()
Recalc
End Sub

Private Sub txtcomm_Change()
Recalc
End Sub

Private Sub txtfare_Change()
Recalc
End Sub

Private Sub txtfee_Change()
Recalc
End Sub

Sub Recalc()

Me.txtvat = CDbl(Me.txtcomm) * 0.1

Me.txttotal = Format((CDbl(Me.txtasf) + CDbl(Me.txtfee) + CDbl(Me.txtinsurance) + CDbl(Me.txtfare)) - CDbl(Me.txtcomm), "###,##.00")
End Sub
 

Private Sub txtinsurance_Change()
Recalc
End Sub

Private Sub txtvat_Change()
Recalc
End Sub
