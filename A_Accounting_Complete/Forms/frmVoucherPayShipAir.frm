VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Begin VB.Form frmVoucherPayShipAir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Pay Ship/Airline"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14805
   Icon            =   "frmVoucherPayShipAir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   60
      TabIndex        =   34
      Top             =   2670
      Width           =   6375
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   960
         TabIndex        =   36
         Top             =   225
         Width           =   1725
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   3375
         TabIndex        =   35
         Top             =   210
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.TextBox txtVoucherAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12075
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   7875
      Width           =   2640
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove"
      Height          =   480
      Left            =   7485
      TabIndex        =   31
      Top             =   8640
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3915
      Left            =   7545
      TabIndex        =   28
      Top             =   3900
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   6906
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Start Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "End Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Particulars"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
      Height          =   645
      Left            =   9825
      TabIndex        =   27
      Top             =   2715
      Width           =   1605
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   5235
      Left            =   0
      TabIndex        =   15
      Top             =   3900
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9234
      SectionData     =   "frmVoucherPayShipAir.frx":030A
   End
   Begin VB.CommandButton cmdmyInsert 
      Caption         =   "Insert"
      Height          =   645
      Left            =   6615
      TabIndex        =   14
      Top             =   2700
      Width           =   1620
   End
   Begin VB.PictureBox Picture3 
      Height          =   2595
      Left            =   45
      ScaleHeight     =   2535
      ScaleWidth      =   14655
      TabIndex        =   2
      Top             =   30
      Width           =   14715
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Voucher Details"
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   10230
         TabIndex        =   18
         Top             =   120
         Width           =   4395
         Begin VB.TextBox txtCheckNo 
            Enabled         =   0   'False
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
            Left            =   2175
            TabIndex        =   24
            Top             =   1230
            Width           =   2175
         End
         Begin VB.TextBox txtTotalAmount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   1755
            Width           =   2175
         End
         Begin VB.ComboBox cboBank 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   195
            Width           =   2220
         End
         Begin VB.ComboBox cboAccount 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   705
            Width           =   2220
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK # :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   885
            TabIndex        =   26
            Top             =   1275
            Width           =   1110
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "SUB TOTAL :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   75
            TabIndex        =   25
            Top             =   1815
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT # :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   555
            TabIndex        =   22
            Top             =   690
            Width           =   1845
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "BANK NAME :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   525
            TabIndex        =   21
            Top             =   240
            Width           =   1590
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Shipping/Airline Below"
         ForeColor       =   &H000000FF&
         Height          =   1275
         Left            =   6540
         TabIndex        =   16
         Top             =   1200
         Width           =   3660
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   135
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   345
            Width           =   3375
         End
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   105
         TabIndex        =   11
         Top             =   600
         Width           =   2460
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   3900
         TabIndex        =   10
         Top             =   570
         Width           =   2460
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   ">"
         Height          =   450
         Left            =   2745
         TabIndex        =   9
         Top             =   870
         Width           =   990
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<"
         Height          =   450
         Left            =   2745
         TabIndex        =   8
         Top             =   1380
         Width           =   990
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         ForeColor       =   &H00FF0000&
         Height          =   1065
         Left            =   6525
         TabIndex        =   3
         Top             =   120
         Width           =   3675
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1305
            TabIndex        =   4
            Top             =   180
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   556
            _Version        =   393216
            Format          =   54263809
            CurrentDate     =   37497
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   1305
            TabIndex        =   5
            Top             =   660
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   556
            _Version        =   393216
            Format          =   54263809
            CurrentDate     =   37497
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   810
            TabIndex        =   7
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            Height          =   255
            Left            =   570
            TabIndex        =   6
            Top             =   195
            Width           =   615
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Ticket Type"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   2460
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Selected Ticket Type"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3900
         TabIndex        =   12
         Top             =   210
         Width           =   2445
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   645
      Left            =   13110
      TabIndex        =   1
      Top             =   2670
      Width           =   1605
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ReCalc"
      Enabled         =   0   'False
      Height          =   645
      Left            =   8220
      TabIndex        =   0
      Top             =   2700
      Width           =   1605
   End
   Begin VB.Label Label5 
      Caption         =   "Voucher Amount :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9795
      TabIndex        =   33
      Top             =   7890
      Width           =   2235
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Sales Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   405
      Left            =   15
      TabIndex        =   30
      Top             =   3435
      Width           =   7425
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Sales Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   405
      Left            =   7575
      TabIndex        =   29
      Top             =   3450
      Width           =   7155
   End
End
Attribute VB_Name = "frmVoucherPayShipAir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAirline           As ADODB.Recordset
Dim SQL                 As String
Dim TempRecord          As String
Dim lngIndex            As Long

Function ExtractIT() As String
Dim i As Integer
Dim TicketType(1 To 3)  As String 'Declare Dynamic Array
Dim TicketTypeFlag      As Boolean
Dim Tmp
Dim myTmpSQL            As String

Dim sDate               As Date
Dim eDate               As Date
Dim MyCriteria          As String

            
                 If Me.List2.ListCount > 0 Then
                 
                                sDate = Format(Me.DTPicker1.Value, "mm/dd/yyyy")
                                eDate = Format(Me.DTPicker2.Value, "mm/dd/yyyy")
                                
                                
                            If Me.List2.ListCount = 1 Then
                                MyCriteria = " where [AirlineName] ='" & Me.Combo1 & "'" & _
                                             " AND [Date] Between #" & sDate & "# AND #" & eDate & "#" & _
                                             " AND [Ticket Type]='" & Me.List2.List(0) & "'" & _
                                             " ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                SQL = "SELECT * FROM qrySalesRptEVAT " & MyCriteria
                            End If
                             
                             If Me.List2.ListCount = 2 Then
                             
                           
                                Call SpoolReport
                                
                                Tmp = kulotRead(App.Path & "\Settings.txt")
                                Select Case CDbl(Tmp)
                                Case 1
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool01"
                                Case 2
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool02"
                                Case 3
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool03"
                                Case 4
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool04"
                                End Select
                                MyCriteria = myTmpSQL & _
                                             " WHERE [Ticket Type]='" & Me.List2.List(0) & "'" & _
                                             " OR [Ticket Type]= '" & Me.List2.List(1) & "'" & _
                                             " ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                SQL = MyCriteria
                          
                             End If
                             
                             If Me.List2.ListCount = 3 Then
                                Call SpoolReport
                                
                                Tmp = kulotRead(App.Path & "\Settings.txt")
                                Select Case CDbl(Tmp)
                                Case 1
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool01"
                                Case 2
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool02"
                                Case 3
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool03"
                                Case 4
                                    myTmpSQL = "SELECT * FROM tbl_Rpt_Sales_Spool04"
                                End Select
                                MyCriteria = myTmpSQL & _
                                             " WHERE [Ticket Type]='" & Me.List2.List(0) & "'" & _
                                             " OR [Ticket Type]= '" & Me.List2.List(1) & "'" & _
                                             " OR [Ticket Type]= '" & Me.List2.List(2) & "'" & _
                                             " ORDER BY [Ticket No], [Ticket Type], [AirlineName]"
                                SQL = MyCriteria
                             
                             End If
                                
                                
                            '==============================================
                           
                                
                           ' Set Me.ARViewer1.ReportSource = Rpt

                            
                            ExtractIT = SQL



                            '==============================================
                  End If

            

End Function

Function IsthereRecord(param) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL As String

Set rst = New ADODB.Recordset
SQL = param
With rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            IsthereRecord = True
        Else
            IsthereRecord = False
        End If
.Close
Set rst = Nothing
End With
End Function


Private Sub cmdDelete_Click()
    Dim intRes As Integer
    intRes = MsgBox("Are you sure you want to delete the selected item?", vbQuestion + vbYesNo, "Delete Confirmation")
    If intRes = vbNo Then Exit Sub
        ListView1.ListItems.Remove (lngIndex)
End Sub

Private Sub cmdmyInsert_Click()
Dim Rpt     As New RptVoucher
Dim ctr     As Integer

ctr = Me.ListView1.ListItems.Count

If CheckVoucher(Me.Combo1, CDate(RetDate(Me.DTPicker1)), CDate(RetDate(Me.DTPicker2))) Then
    MsgBox "This Sales was already Vouchered!", vbInformation
    Exit Sub
End If

If ctr > 0 Then
    If Me.ListView1.ListItems(ctr).Text = Format(Me.DTPicker1, "mm/dd/yyyy") And _
            Me.ListView1.ListItems(ctr).SubItems(1) = Format(Me.DTPicker2, "mm/dd/yyyy") Then
            MsgBox "This was already inserted", vbInformation
            Exit Sub
    End If
End If

If Me.List2.ListCount > 0 Then
    With Rpt
      .DataControl1.Connection = cn
      .DataControl1.Source = ExtractIT
      .lblVoucherAmount = "0.00"
    End With

    If IsthereRecord(ExtractIT) Then
    '    Call InsertToLSTVW
    Set Me.ARViewer1.ReportSource = Rpt
    Me.cmdRefresh.Enabled = True
    Else
        MsgBox "No Sales for this period :[" & RetDate(Me.DTPicker1) & "]-[" & RetDate(Me.DTPicker2) & "]", vbCritical
    End If
 Else
    MsgBox "Please select ticket type", vbInformation
End If
End Sub

Sub InsertToLSTVW()
Dim myList  As ListItem
Set myList = ListView1.ListItems.Add(, , Format(Me.DTPicker1, "mm/dd/yyyy"))
    myList.SubItems(1) = Format(Me.DTPicker2, "mm/dd/yyyy")
    myList.SubItems(2) = Me.Combo1
    myList.SubItems(3) = Format(Me.txtTotalAmount, "###,##0.00")
'Call cmdRefresh_Click
End Sub

Private Sub cmdPost_Click()
On Error GoTo FailSafe_Err
Dim rst             As New ADODB.Recordset
Dim SQL             As String
Dim i               As Integer
Dim myParticular    As String
Dim ask             As Integer

SQL = "SELECT * FROM tbl_Voucher ORDER by VoucherID ASC"

If ZeroBal Then
    MsgBox "cannot continue insufficient amount for current account..." & Me.cboAccount, vbInformation
    Exit Sub
End If
If CheckNull(Me.cboBank) Then: MsgBox "Please select Bank!", vbInformation: Exit Sub
If CheckNull(Me.cboAccount) Then: MsgBox "Please select account!", vbInformation: Exit Sub
If Me.txtCheckNo = "No Checks Available!" Then: MsgBox "Invalid check!", vbInformation: Exit Sub

ask = MsgBox("Sure to save this?", vbInformation + vbYesNo)
If ask = vbNo Then: Exit Sub
With rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
         cn.BeginTrans
        .AddNew
                .Fields("Payto").Value = Me.Combo1
                .Fields("Date").Value = RetDate(Now)
                .Fields("BankID").Value = FindBankID(Me.cboBank)
                .Fields("CheckNo").Value = Me.txtCheckNo
                .Fields("TotalAmount").Value = Me.txtVoucherAmount
                .Fields("Has Issued Check").Value = True
                .Fields("Cash").Value = False
                .Fields("Check").Value = True
                .Fields("For Refund").Value = False
                .Fields("Electronic Transfer").Value = False
        .Update
        
        If Me.ListView1.ListItems.Count > 0 Then
                For i = 1 To Me.ListView1.ListItems.Count
                myParticular = Me.ListView1.ListItems(i).Text & "-" & Me.ListView1.ListItems(i).SubItems(1) & " " & Me.ListView1.ListItems(i).SubItems(2)
                        Call VoucherDetailsAdd(.Fields("VoucherID").Value, myParticular, Me.ListView1.ListItems(i).SubItems(3), Me.ListView1.ListItems(i).Text, Me.ListView1.ListItems(i).SubItems(1))
                Next i
        End If
        
        
'// Now Deduct this voucher to bank
    Call UpdatePassbook(CDbl(Me.txtVoucherAmount), _
    IIf(Me.Option2 = True, Me.txtCheckNo, ""), _
    Format(Now, "mm/dd/yyyy"), _
    "Issued Voucher to :" & Me.Combo1 & " as payment(s) for sales from " & _
                            RetDate(Me.DTPicker1) & " to " & RetDate(Me.DTPicker2), _
    Me.cboAccount, "n/a", _
    IIf(Me.Option1 = True, CDbl(Me.txtVoucherAmount), 0), 0, _
    IIf(Me.Option2 = True, CDbl(Me.txtVoucherAmount), 0), 0, "", "", "", "", "", "", "", .Fields("VoucherID").Value)
    Call UpdateCheck
        
        
        cn.CommitTrans
        MsgBox "Voucher Save", vbInformation
       .Close
     Set rst = Nothing
End With
Exit Sub
FailSafe_Err:
cn.RollbackTrans
MsgBox "There was an error while saving the voucher", vbInformation

End Sub

Function CheckVoucher(usrPayto, usrDate1, usrDate2) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM tbl_VoucherDetails WHERE [Payto]='" & usrPayto & "' AND [Date1]=#" & CDate(usrDate1) & "# AND [Date2]=#" & CDate(usrDate2) & "#"
With rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                CheckVoucher = True
                     Else
                CheckVoucher = False
            End If
       .Close
     Set rst = Nothing
End With

End Function
Sub VoucherDetailsAdd(param, usrParticulars, usrAmount, usrDate1, usrDate2)
Dim rst As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT * FROM tbl_VoucherDetails"
With rst
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
        .AddNew
                .Fields("VoucherID").Value = param
                .Fields("Particulars").Value = usrParticulars
                .Fields("Amount").Value = CDbl(usrAmount)
                .Fields("Payto").Value = Me.Combo1
                .Fields("Date1").Value = CDate(RetDate(usrDate1))
                .Fields("Date2").Value = CDate(RetDate(usrDate2))
        .Update
       .Close
     Set rst = Nothing
End With


End Sub

Private Sub cmdRefresh_Click()
Dim Rpt  As New RptVoucher
 
 With Rpt
   .DataControl1.Connection = cn
   .DataControl1.Source = ExtractIT
   myglobal_NetSales = myGlobal_CashDue - myGlobal_RefundAmt
   .lblVoucherAmount = Format(myglobal_NetSales, "###,##0.00")
   Me.txtTotalAmount = RetCurrency(myglobal_NetSales)
   
 End With

 Set Me.ARViewer1.ReportSource = Rpt
 Call InsertToLSTVW
 Me.txtVoucherAmount = RetCurrency(SumListView)
 Me.cmdRefresh.Enabled = False
End Sub

Function SumListView() As Double
Dim y As Integer
Dim Tmp As Double

If Me.ListView1.ListItems.Count > 0 Then
        For y = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(y).SubItems(3) <> "" Then
                Tmp = Tmp + CDbl(Me.ListView1.ListItems(y).SubItems(3))
            End If
        Next y
Else
        Tmp = 0
End If
SumListView = Tmp
End Function


Private Sub cboAccount_Click()
Me.txtCheckNo = ""
'Return Check if check is selected
Me.txtCheckNo = ReturnCheck()
End Sub

Private Sub cboBank_Click()
Call FillAccount
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdInsert_Click()
If Len(Me.List1.List(Me.List1.ListIndex)) > 0 Then
Me.List2.AddItem Me.List1.List(Me.List1.ListIndex)
Me.List1.RemoveItem (Me.List1.ListIndex)
End If
End Sub


Private Sub cmdRemove_Click()
If Len(Me.List2.List(Me.List2.ListIndex)) > 0 Then
Me.List1.AddItem Me.List2.List(Me.List2.ListIndex)
Me.List2.RemoveItem (Me.List2.ListIndex)
End If

End Sub

Private Sub Combo1_Click()
Call FillList
End Sub



Private Sub DTPicker1_Change()
    If DTPicker1.Value > DTPicker2.Value Then
        DTPicker2.Value = DTPicker1.Value
    End If
End Sub

Private Sub DTPicker2_Change()
    If DTPicker2.Value > DTPicker1.Value Then
        DTPicker1.Value = DTPicker2.Value
    End If
End Sub

Private Sub Form_Activate()

    
    Me.List2.Clear
    
    Call FillBank
    Call FillCombo
    Call FillList
    
End Sub

Private Sub Form_Load()


    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    
    
End Sub


Sub FillCombo()
Me.Combo1.Clear
Set RsAirline = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]<>'NONE'"

With RsAirline
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
            Me.Combo1.AddItem .Fields(1).Value
            .MoveNext
        Loop
        Me.Combo1.ListIndex = 0
    End If
End With
End Sub


Sub FillList()
Dim Tmp As ADODB.Recordset
Dim tmpAirline As New ADODB.Recordset
Dim SQL As String

Set tmpAirline = New ADODB.Recordset

SQL = "SELECT * FROM tbl_Airline WHERE [AirlineName]='" & Me.Combo1 & "'"
tmpAirline.Open SQL, cn, adOpenKeyset, adLockOptimistic
If tmpAirline.RecordCount > 0 Then
    TempRecord = tmpAirline.Fields(2).Value
End If


Set Tmp = New ADODB.Recordset

SQL = "SELECT * FROM tbl_TicketType WHERE [AirlineShippingLine]='" & TempRecord & "' ORDER BY [Ticket Type] ASC"
Me.List1.Clear
With Tmp
    .Open SQL, cn, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        Do While Not .EOF
        Me.List1.AddItem .Fields(1).Value
        .MoveNext
        Loop
        Me.List1.ListIndex = 0
    End If
End With

End Sub



Sub FillAccount()
Dim rst As New ADODB.Recordset
SQL = "SELECT DISTINCT  [Account Number] FROM tbl_AccountsSetting WHERE [BankID]=" & FindBankID(Me.cboBank) & " ORDER by [Account Number] ASC"
With rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Me.cboAccount.Clear
                    .MoveFirst
               Do While Not .EOF
                    Me.cboAccount.AddItem .Fields("Account Number").Value
                    .MoveNext
               Loop
        Else
            Me.cboAccount.Clear
            Me.txtCheckNo = ""
        End If
End With

End Sub

Sub FillBank()
Dim rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks"
        Me.cboBank.Clear
With rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Do While Not .EOF
                Me.cboBank.AddItem .Fields(1).Value
                .MoveNext
            Loop
        End If
       .Close
     Set rst = Nothing
End With
End Sub

Function ReturnCheck() As String
Dim Tmp As ADODB.Recordset
Set Tmp = New ADODB.Recordset
SQL = "SELECT  * FROM tbl_checks WHERE [BankID]=" & FindBankID(Me.cboBank) & " AND [Status]='Un-Used' AND [AccNo]='" & Me.cboAccount & "'ORDER by [CheckNo] ASC"
With Tmp
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReturnCheck = .Fields("CheckNo").Value
            
        Else
            ReturnCheck = "No Checks Available!"
        End If
      .Close
    Set Tmp = Nothing
End With

End Function

Function FindBankID(param) As Long
Dim rst As New ADODB.Recordset
SQL = "SELECT * FROM tbl_Banks WHERE [Bank Name]='" & UCase(param) & "'"
With rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
              FindBankID = .Fields(0).Value
          Else
              FindBankID = -1
        End If
     .Close
   Set rst = Nothing
End With
End Function

Private Sub ListView1_Click()
If ListView1.ListItems.Count > 0 Then
        lngIndex = ListView1.SelectedItem.Index
End If
End Sub

Private Sub Option1_Click()
frmPO_DomesticFind.Tag = "voucher"
frmPO_DomesticFind.Show 1
End Sub

Sub SpoolReport()
'On Error GoTo FailSafe_Error
Dim mySQL               As String
Dim recs                As Long
Dim Tmp
Dim myInsertSQL         As String
Dim myInsertSQL_Head    As String

Tmp = kulotRead(App.Path & "\Settings.txt")




mySQL = "(StatementDetails, TransID, [Date], [Ticket No]," & _
                 "Name, nVoid, Route, misc, nMisc, nGross, [nCompany Commission]," & _
                 "xVAT, nVAT, nSumOfInsurance, Insurance, ASF, nSumOfASF," & _
                 "[nTerminal Fee], [nSumOfTerminal Fee], nMeals," & _
                 "Comm, qVAT, [Ticket Type], AirlineID, AirlineName,"

mySQL = mySQL & "TotIns, qEvat, Subtotal, CashDue, nFare, Profit, Fare, Gross," & _
                "[Company Commission], SumOfInsurance, SumOfASF, [Terminal Fee], " & _
                "[SumOfTerminal Fee], Meals)"

mySQL = mySQL & "SELECT tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID," & _
                "tbl_Statement.Date, tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name," & _
                "tbl_StatementDetail.Void AS nVoid, tbl_StatementTickets.Route, qSum_MISC.misc," & _
                " IIf([nVoid]=False,[misc],0) AS nMisc, IIf([nVoid]=False,[Gross],0) AS nGross," & _
                " IIf([nVoid]=False,[Company Commission],0) AS [nCompany Commission], " & _
                "tbl_StatementTickets.VAT AS xVAT,"
      
mySQL = mySQL & "IIf([nVoid]=False,IIf(IsNull([xVat]),0,[xVat]),0) AS nVAT," & _
                " IIf([nVoid]=False,[SumOfInsurance],0) AS nSumOfInsurance," & _
                " tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF," & _
                " IIf([nVoid]=False,[SumOfASF],0) AS nSumOfASF, IIf([nVoid]=False," & _
                "[Terminal Fee],0) AS [nTerminal Fee], IIf([nVoid]=False,[SumOfTerminal Fee],0)" & _
                " AS [nSumOfTerminal Fee], IIf([nVoid]=False,[Meals],0) AS nMeals, " & _
                "IIf([nVoid]=False,Format([nGross]*([nCompany Commission])/100,'Standard'),0) AS Comm,"
      
mySQL = mySQL & "IIf([nVoid]=False,[Comm]*([nVat]/100),0) AS qVAT, tbl_StatementDetail.[Ticket Type]," & _
                " tbl_Airline.AirlineID, tbl_Airline.AirlineName, IIf([nVoid]=False,[nSumOfInsurance],0)" & _
                " AS TotIns, IIf([nVoid]=False,([nGross]+[TotIns]+[nSumOfASF]+[nSumOfTerminal Fee]" & _
                "+[nMeals]+[nMisc])*0.12,0) AS qEvat, IIf([nVoid]=False,[nGross]-[Comm]+[TotIns]" & _
                "+[nSumOfASF]+[nSumOfTerminal Fee]+[nMeals]+[qvat]+[nMisc],0) AS Subtotal, " & _
                "IIf([nVoid]=False,IIf([AirlineID]=36,Abs(Int(([Subtotal]+[qEvat])*-1))," & _
                "[Subtotal]+[qEvat]),0) AS CashDue, IIf([nvoid]=False,[Fare],0) AS nFare, " & _
                "[nFare]-IIf(IsNull([Cashdue]),0,[CashDue]) AS Profit, tbl_StatementDetail.Fare," & _
                "tbl_StatementDetail.Gross, tbl_StatementTickets.[Company Commission]," & _
                "q1.SumOfInsurance, qSum_ASF.SumOfASF, tbl_StatementTickets.[Terminal Fee], " & _
                "qSum_TF.[SumOfTerminal Fee], tbl_StatementTickets.Meals"
mySQL = mySQL & " FROM tbl_Statement INNER JOIN ((((((tbl_Airline RIGHT JOIN tbl_StatementDetail" & _
                " ON tbl_Airline.AirlineID=tbl_StatementDetail.Airline) LEFT JOIN q1 " & _
                " ON tbl_StatementDetail.StatementDetails=q1.StatementDetails) " & _
                " LEFT JOIN qSum_ASF ON tbl_StatementDetail.StatementDetails=qSum_ASF.StatementDetails) " & _
                " LEFT JOIN qSum_TF ON tbl_StatementDetail.StatementDetails=qSum_TF.StatementDetails) " & _
                " LEFT JOIN qSum_MISC ON tbl_StatementDetail.StatementDetails=qSum_MISC.StatementDetails)" & _
                " INNER JOIN tbl_StatementTickets ON tbl_StatementDetail.StatementDetails " & _
                "=tbl_StatementTickets.StatementDetails) ON tbl_Statement.TransID=tbl_StatementDetail.TransID" & _
                " GROUP BY tbl_StatementDetail.StatementDetails, tbl_StatementDetail.TransID, " & _
                "tbl_Statement.Date, tbl_StatementDetail.[Ticket No], tbl_StatementDetail.Name, " & _
                "tbl_StatementDetail.Void, tbl_StatementTickets.Route, qSum_MISC.misc, " & _
                "tbl_StatementTickets.VAT, tbl_StatementTickets.Insurance, tbl_StatementTickets.ASF," & _
                " tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineID, tbl_Airline.AirlineName, " & _
                "tbl_StatementDetail.Fare, tbl_StatementDetail.Gross, tbl_StatementTickets.[Company Commission], q1.SumOfInsurance," & _
                " qSum_ASF.SumOfASF, tbl_StatementTickets.[Terminal Fee], qSum_TF.[SumOfTerminal Fee], tbl_StatementTickets.Meals " & _
                " HAVING (((tbl_Statement.Date) Between #" & Format(Me.DTPicker1, "MM/DD/YYYY") & "# And #" & Format(Me.DTPicker2, "MM/DD/YYYY") & "#) AND ((tbl_Airline.AirlineName)='" & Me.Combo1 & "'))" & _
                " ORDER BY tbl_StatementDetail.[Ticket No], tbl_StatementDetail.[Ticket Type], tbl_Airline.AirlineName;"



Select Case CDbl(Tmp)
Case 1
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool01"
        mySQL = myInsertSQL_Head & mySQL
Case 2
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool02"
        mySQL = myInsertSQL_Head & mySQL

Case 3
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool03"
        mySQL = myInsertSQL_Head & mySQL

Case 4
        myInsertSQL_Head = "INSERT INTO tbl_Rpt_Sales_Spool04"
        mySQL = myInsertSQL_Head & mySQL

End Select

cn.BeginTrans
        'Remove first existing Records
         Select Case CDbl(Tmp)
            Case 1
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool01"
            Case 2
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool02"
            Case 3
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool03"
            Case 4
                cn.Execute "DELETE * FROM tbl_Rpt_Sales_Spool04"
         End Select
        

        'Now append to spool table
        cn.Execute mySQL, recs, adExecuteNoRecords
cn.CommitTrans
Exit Sub

FailSafe_Error:
    cn.RollbackTrans
End Sub

Sub UpdatePassbook(ByVal nAmt As Double, _
    ByVal CheckNo As String, ByVal CheckDate As String, _
    Optional Desc As String, Optional ByVal AccNo As String, _
    Optional strAir As String, _
    Optional nCash, Optional nCard, Optional nCheck, _
    Optional nOthers, Optional nCardName, _
    Optional nCardNumber, Optional nCardHolder, _
    Optional nBank1, Optional nBank2, _
    Optional nBank3, Optional nBank4, Optional usrVID)
    
'On Error Resume Next
Dim RsPassbk As New ADODB.Recordset
Dim RsBalance As New ADODB.Recordset
Dim TempBal As Double

SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & AccNo & "'"
With RsPassbk
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    .MoveFirst
                 TempBal = .Fields("Current Balance").Value
            End If
            .Close
      Set RsPassbk = Nothing
End With




SQL = "SELECT * FROM tbl_BankPassbook"
With RsPassbk
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If IsExist_Voucher(usrVID) Then
            .MoveFirst
            .Find "[Voucher No]='" & usrVID & "'"
        Else
            .AddNew
        End If
            .Fields("Deposit Date").Value = Format(Now, "mm/dd/yyyy")
            .Fields("Check No").Value = CheckNo
            .Fields("Check Date").Value = CheckDate
            .Fields("Voucher No").Value = usrVID
            .Fields("Description").Value = Desc
            .Fields("Credit").Value = 0
            .Fields("Debit").Value = nAmt
            .Fields("Account Number").Value = AccNo
            .Fields("Balance").Value = TempBal - nAmt
            .Fields("Cash Amount").Value = CDbl(nCash)
            .Fields("Card Amount").Value = CDbl(nCard)
            .Fields("Check Amount").Value = CDbl(nCheck)
            .Fields("Others Amount").Value = CDbl(nOthers)
            .Fields("ORno").Value = "n/a"
            .Fields("Airline").Value = -1   'strAir
            .Fields("Card Name").Value = nCardName
            .Fields("Card Number").Value = nCardNumber
            .Fields("Card Holder").Value = nCardHolder
            .Fields("Bank1").Value = nBank1
            .Fields("Bank2").Value = nBank2
            .Fields("Bank3").Value = nBank3
            .Fields("Bank4").Value = nBank4
        
            .Update
End With

SQL = "UPDATE tbl_AccountsSetting SET [Current Balance] = " & _
              CDbl(TempBal - nAmt) & " WHERE [Account Number]= '" & UCase(AccNo) & "'"
              cn.BeginTrans
                    cn.Execute SQL
              cn.CommitTrans
Exit Sub
FailSafe_Error:
cn.RollbackTrans
End Sub

Function ZeroBal() As Boolean
Dim RsPassbk As New ADODB.Recordset
Dim RsBalance As New ADODB.Recordset
Dim TempBal As Double

SQL = "SELECT * FROM tbl_AccountsSetting WHERE [Account Number]='" & Me.cboAccount & "'"
With RsPassbk
            .Open SQL, cn, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                    .MoveFirst
                 TempBal = .Fields("Current Balance").Value
                 
                 If CDbl(TempBal) <= 0 Then
                     ZeroBal = True
                 Else
                     ZeroBal = False
                 End If
            End If
            .Close
      Set RsPassbk = Nothing
End With
End Function


Sub UpdateCheck()
cn.BeginTrans
SQL = "UPDATE tbl_checks SET [Status] = 'issued' WHERE [CheckNo]='" & Me.txtCheckNo & "' AND [AccNo]='" & Me.cboAccount & "' AND [Status]='Un-Used' "
cn.Execute SQL
cn.CommitTrans
End Sub


Function IsExist_Voucher(param) As Boolean
Dim rst         As New ADODB.Recordset
Dim SQL         As String

SQL = "SELECT * FROM qryBankPassbook WHERE [Voucher No]='" & param & "'"
With rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                    IsExist_Voucher = True
            Else
                    IsExist_Voucher = False
        End If
        .Close
      Set rst = Nothing
End With
End Function


