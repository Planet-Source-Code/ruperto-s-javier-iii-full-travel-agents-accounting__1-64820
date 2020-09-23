VERSION 5.00
Begin VB.Form frmFareBasisTmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Amount"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   Icon            =   "frmFareBasisTmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   3615
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1485
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   4
         Top             =   210
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1485
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   690
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Fare Basis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   765
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   1380
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   2310
      TabIndex        =   0
      Top             =   1410
      Width           =   1350
   End
End
Attribute VB_Name = "frmFareBasisTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsFareBasis As ADODB.Recordset
Dim SQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo FailSafe_Error
If CDbl(Me.Text2) = 0 Then
            MsgBox "Amount should be greater zero", vbCritical, "ELS"
                Call kulotHL(Me.Text2)
            Exit Sub
End If

If Not IsNumeric(Me.Text2) Then
            MsgBox "Amount requires numeric value", vbCritical, "ELS"
                Call kulotHL(Me.Text2)
            Exit Sub
End If
                With RsFareBasis
                        .Find "[FareBasisID]=" & Me.Tag
                    If Not .EOF Or Not .BOF Then
                        .Fields(1).Value = UCase(Me.Text1)
                        .Fields(2).Value = CDbl(Me.Text2)
                        .Update
                        MsgBox "Price Updated...", vbInformation, "ELS"
                    End If
                End With
frmSetTicketPricing.DisplayFB
Exit Sub
FailSafe_Error:
                        MsgBox "There was an error in saving the price", vbInformation, "ELS"
End Sub

Private Sub Form_Load()
SQL = "SELECT * FROM tbl_TmpFareBasis"
Set RsFareBasis = New ADODB.Recordset
RsFareBasis.Open SQL, cn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Text2_GotFocus()
    Call kulotHL(Me.Text2)
End Sub
