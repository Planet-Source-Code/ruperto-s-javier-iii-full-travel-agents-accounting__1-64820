VERSION 5.00
Begin VB.Form frmEditRoutes 
   Caption         =   "Edit Routes"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   Icon            =   "frmEditRoutes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRouteID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1785
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   90
      Width           =   2325
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   465
      Left            =   2535
      TabIndex        =   5
      Top             =   1425
      Width           =   1605
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   945
      TabIndex        =   4
      Top             =   1425
      Width           =   1605
   End
   Begin VB.TextBox txtOrigin 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   540
      Width           =   2310
   End
   Begin VB.TextBox txtDestination 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   945
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ROUTE ID :"
      Height          =   285
      Left            =   300
      TabIndex        =   7
      Top             =   135
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PLACE OF ORIGIN :"
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Top             =   555
      Width           =   1890
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINATION :"
      Height          =   285
      Left            =   570
      TabIndex        =   2
      Top             =   960
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditRoutes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo FailSafe_Error
Dim ask As Integer
SQL = "UPDATE tbl_Routes SET [From] = '" & UCase(Me.txtOrigin) & "', [To] = '" & UCase(Me.txtDestination) & "' WHERE (((RouteID)=" & CLng(Me.txtRouteID) & "))"
ask = MsgBox("Are you sure you want to save?", vbInformation + vbYesNo)
If ask = vbYes Then
cn.BeginTrans
cn.Execute SQL
cn.CommitTrans
            MsgBox "Record Save...", vbInformation
            Unload Me
End If
Exit Sub
FailSafe_Error:
cn.RollbackTrans
End Sub

Private Sub Form_Load()
Me.txtRouteID = frmAddRoutes.DataGrid1.Columns(0).Text
Me.txtOrigin = frmAddRoutes.DataGrid1.Columns(1).Text
Me.txtDestination = frmAddRoutes.DataGrid1.Columns(2).Text
End Sub

