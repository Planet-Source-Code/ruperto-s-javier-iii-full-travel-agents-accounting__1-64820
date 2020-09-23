VERSION 5.00
Begin VB.Form frmSetCOMM 
   Caption         =   "Set COMMISSION"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   Icon            =   "frmSetCOMM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   555
      Left            =   2490
      TabIndex        =   4
      Top             =   2775
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4155
      TabIndex        =   3
      Top             =   2775
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set COMMISSION for Shipping/Airline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   5790
      Begin VB.PictureBox Picture1 
         Height          =   885
         Left            =   75
         ScaleHeight     =   825
         ScaleWidth      =   5535
         TabIndex        =   9
         Top             =   690
         Width           =   5595
         Begin VB.ComboBox cboFrom 
            Height          =   315
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   60
            Width           =   3285
         End
         Begin VB.ComboBox cboTo 
            Height          =   315
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   3285
         End
         Begin VB.Label Label4 
            Caption         =   "From (Route)"
            Height          =   300
            Left            =   120
            TabIndex        =   13
            Top             =   75
            Width           =   1230
         End
         Begin VB.Label Label5 
            Caption         =   "To (Destination)"
            Height          =   300
            Left            =   120
            TabIndex        =   12
            Top             =   435
            Width           =   1230
         End
      End
      Begin VB.TextBox txtVat 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2310
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   2175
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2310
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1740
         Width           =   3300
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3285
      End
      Begin VB.Label Label3 
         Caption         =   "VAT %"
         Height          =   285
         Left            =   165
         TabIndex        =   8
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "COMMISSION %"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   1785
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Shippling/Airline Name :"
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmSetCOMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsComm As ADODB.Recordset
Dim SQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSet_Click()
If CheckNull(Me.cboFrom) And CheckNull(Me.cboTo) Then
    MsgBox "Please select from and to route", vbInformation
Else
    Call CheckStat
End If

End Sub

Private Sub Form_Load()
Call FillCombo
Me.Combo1.ListIndex = 0
Call FillRoutes(1)
Call FillRoutes(2)
End Sub


Sub FillRoutes(Param)
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
If Param = 1 Then
SQL = "SELECT DISTINCT [From] From qryRoutePricing ORDER BY [From]"
Else
SQL = "SELECT DISTINCT [TO] From qryRoutePricing ORDER BY [TO]"
End If
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
   If Param = 1 Then
        Me.cboTo.Clear
        Do While Not .EOF
            Me.cboFrom.AddItem .Fields(0).Value
        .MoveNext
        Loop
     Else
        Me.cboTo.Clear
        Do While Not .EOF
            Me.cboTo.AddItem .Fields(0).Value
        .MoveNext
        Loop
     
   End If
    End If
End With

End Sub

Sub CommUpdate(ByVal strShip As String, ByVal cAmt As Double, ByVal nVat)
'On Error GoTo FailSafe_Error
SQL = "UPDATE tbl_RoutePricing SET [Commision] = " & cAmt & " WHERE [AirlineID]= " & FindAirline(strShip) & " AND [RouteID]= " & ReturnRouteID(Me.cboFrom, Me.cboTo)
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans

SQL = "UPDATE tbl_Computation SET tbl_Computation.Commission = " & cAmt & " , tbl_Computation.VAT =" & CDbl(nVat)
cn.BeginTrans
    cn.Execute SQL
cn.CommitTrans

MsgBox "Commission for " & Me.Combo1 & " successfully updated!!!", vbInformation
Exit Sub
FailSafe_Error:
cn.RollbackTrans
MsgBox "Error updating for " & Me.Combo1 & " please try again...", vbInformation
End Sub


Function ReturnRouteID(ByVal strFrom As String, ByVal strTo As String) As Long
Dim Rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Routes WHERE [From]= '" & strFrom & "' AND [To] ='" & strTo & "'"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
       If .RecordCount > 0 Then
            ReturnRouteID = .Fields(0).Value
       Else
            ReturnRouteID = -1
       End If
       .Close
    Set Rst = Nothing
End With
End Function


Sub CheckStat()
Dim Rst As New ADODB.Recordset
Dim ask As Integer
SQL = "SELECT * FROM tbl_RoutePricing WHERE [AirlineID]= " & FindAirline(Me.Combo1) & " AND [RouteID]= " & ReturnRouteID(Me.cboFrom, Me.cboTo)
With Rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
               ask = MsgBox("You are about to update " & .RecordCount & " record(s) continue?", vbInformation + vbYesNo)
               If ask = vbYes Then
                Call CommUpdate(Me.Combo1, CDbl(Me.Text1), Me.txtVat)
               End If
            Else
                MsgBox "There are no records that match your selection", vbInformation
            End If
End With
End Sub

Sub FillCombo()
Dim Tmp As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline"
With Tmp
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
                Me.Combo1.Clear
                Do While Not .EOF
                    Me.Combo1.AddItem .Fields(1).Value
                .MoveNext
                Loop
           End If
End With
End Sub

Function FindAirline(Param) As Long
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Param & "'"
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        FindAirline = .Fields(0).Value
      Else
        FindAirline = -1
    End If
    .Close
End With
Set Tmp = Nothing
End Function

Private Sub Text1_Change()
If Not IsNumeric(Me.Text1) Then
            Me.Text1 = "0.00"
            Call kulotHL(Me.Text1)
End If
End Sub
