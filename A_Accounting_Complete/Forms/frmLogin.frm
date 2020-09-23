VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "lvbuttons.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Accounting.ucGradContainer ucGradContainer1 
      Height          =   3120
      Left            =   0
      TabIndex        =   3
      Top             =   -15
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   5503
      Caption         =   ""
      CaptionAlignment=   2
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ImageList SmallImages 
         Left            =   6345
         Top             =   315
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   35
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":0CDA
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":1B2C
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":297E
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":3258
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":3B32
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":440C
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":4DD6
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":56B0
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":59CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":62A4
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":6B7E
               Key             =   "IMG12"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":7458
               Key             =   "IMG13"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":7772
               Key             =   "IMG14"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":804C
               Key             =   "IMG15"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":8926
               Key             =   "IMG16"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":9200
               Key             =   "IMG17"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":9ADA
               Key             =   "IMG18"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":A3B4
               Key             =   "IMG19"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":AC8E
               Key             =   "IMG20"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":B568
               Key             =   "IMG21"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":BE42
               Key             =   "IMG22"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":C71C
               Key             =   "IMG23"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":CFF6
               Key             =   "IMG24"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":D8D0
               Key             =   "IMG25"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":E1AA
               Key             =   "IMG26"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":EA84
               Key             =   "IMG27"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":F35E
               Key             =   "IMG28"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":FC38
               Key             =   "IMG29"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":10512
               Key             =   "IMG30"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":10DC8
               Key             =   "IMG31"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":116A2
               Key             =   "IMG32"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":11AF4
               Key             =   "IMG33"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":11F46
               Key             =   "IMG34"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogin.frx":146F8
               Key             =   "IMG35"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5130
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2055
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5130
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1575
         Width           =   2430
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FCF4EF&
         Caption         =   "Unmask the Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3810
         TabIndex        =   4
         Top             =   1125
         Visible         =   0   'False
         Width           =   2625
      End
      Begin LVbuttons.LaVolpeButton cmdLogin 
         Height          =   435
         Left            =   3795
         TabIndex        =   2
         Top             =   2550
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Login"
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
         MICON           =   "frmLogin.frx":1597A
         ALIGN           =   1
         IMGLST          =   "SmallImages"
         IMGICON         =   "12"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdExit 
         Height          =   435
         Left            =   5685
         TabIndex        =   5
         Top             =   2550
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   767
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
         MICON           =   "frmLogin.frx":15996
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
      Begin VB.Image Image1 
         Height          =   1830
         Left            =   75
         Picture         =   "frmLogin.frx":159B2
         Stretch         =   -1  'True
         Top             =   1170
         Width           =   3555
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3810
         TabIndex        =   8
         Top             =   2055
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3810
         TabIndex        =   7
         Top             =   1575
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please Supply Valid User name and password for you to use the services in the system..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   570
         Width           =   4500
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim SQL As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text1.PasswordChar = ""
Else
    Text1.PasswordChar = "*"
End If

End Sub

Private Sub cmdExit_Click()
End
End Sub


Private Sub cmdLogin_Click()

Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Users WHERE [UserName]='" & Me.Combo1.Text & "' AND [UserPass]='" & Encrypt(Me.Text1) & "'"
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        Call enableDisable
              With MDImain.StatusBar1
                    .Panels(1).Text = Format(Now, "mm/dd/yyyy")
                    .Panels(2).Text = Rs.Fields(1).Value
                    .Panels(3).Text = Rs.Fields(3).Value
              End With
                    Unload Me
        Else
                MsgBox "Invalid Password!!!"
        End If
        .Close
     Set Rs = Nothing
End With
End Sub

Sub enableDisable()
Dim Forms(1 To 17) As String

Forms(1) = "frmShipAirline"
Forms(2) = "frmAddRoutes"
Forms(3) = "frmAddTicketType"
Forms(4) = "frmAddTickets"
Forms(5) = "frmAddPassengerType"
Forms(6) = "frmSetTicketPricing"
Forms(7) = "frmBankAccSettings"
Forms(8) = "frmAddChecks"
Forms(9) = "frmCustomerAccounts"
Forms(10) = "frmStatement"
Forms(11) = "frmStatementInter"
Forms(12) = "frmExchangeDoc"
Forms(13) = "frmVoidTicket"
Forms(14) = "frmPassporting"
Forms(15) = "frmCashier"
Forms(16) = "frmCashVoucher"
Forms(17) = "frmRefund"

MDImain.mnuf1.Enabled = False
MDImain.mnuRoutes.Enabled = False
MDImain.mnuf3.Enabled = False
MDImain.mnuF5.Enabled = False
MDImain.mnuf7.Enabled = False
MDImain.mnuf9.Enabled = False
MDImain.mnuBankSet.Enabled = False
MDImain.mnuAddChecks.Enabled = False
MDImain.mnuCust.Enabled = False
MDImain.mnuDomestic.Enabled = False
MDImain.mnuInternational.Enabled = False
MDImain.mnuExchangeDoc.Enabled = False
MDImain.mnuvoid.Enabled = False
MDImain.mnuPassporting.Enabled = False
MDImain.mnuCashierPayments.Enabled = False
MDImain.mnuCashierVoucher.Enabled = False
MDImain.mnuTrans6.Enabled = False

With Rs
       If .Fields("form1").Value = Forms(1) Then
            MDImain.mnuf1.Enabled = True
       End If
       
       If .Fields("form2").Value = Forms(2) Then
            MDImain.mnuRoutes.Enabled = True
       End If
       
       If .Fields("form3").Value = Forms(3) Then
MDImain.mnuf3.Enabled = True
       End If
       
       If .Fields("form4").Value = Forms(4) Then
MDImain.mnuF5.Enabled = True
       End If
       
       If .Fields("form5").Value = Forms(5) Then
MDImain.mnuf7.Enabled = True
       End If
       
       If .Fields("form6").Value = Forms(6) Then
MDImain.mnuf9.Enabled = True
       End If
       
       If .Fields("form7").Value = Forms(7) Then
MDImain.mnuBankSet.Enabled = True
       End If
       
       If .Fields("form8").Value = Forms(8) Then
MDImain.mnuAddChecks.Enabled = True
       End If
       
       If .Fields("form9").Value = Forms(9) Then
MDImain.mnuCust.Enabled = True
       End If
       
       If .Fields("form10").Value = Forms(10) Then
MDImain.mnuDomestic.Enabled = True
       End If
       
       If .Fields("form11").Value = Forms(11) Then
MDImain.mnuInternational.Enabled = True
       End If
       
       If .Fields("form12").Value = Forms(12) Then
MDImain.mnuExchangeDoc.Enabled = True
       End If
       
       If .Fields("form13").Value = Forms(13) Then
MDImain.mnuvoid.Enabled = True
       End If
       
       If .Fields("form14").Value = Forms(14) Then
MDImain.mnuPassporting.Enabled = True
       End If
       
       If .Fields("form15").Value = Forms(15) Then
MDImain.mnuCashierPayments.Enabled = True
       End If
       
       If .Fields("form16").Value = Forms(16) Then
MDImain.mnuCashierVoucher.Enabled = True
       End If
       
       If .Fields("form17").Value = Forms(17) Then
MDImain.mnuTrans6.Enabled = True
       End If
       
End With


End Sub

Private Sub Combo1_Click()
SendKeys "{Tab}"
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
TrapEnter (KeyCode)
End Sub


Private Sub Form_Load()
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Users ORDER by [UserName] ASC"
With Rs
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Me.Combo1.Clear
                Do While Not .EOF
                  Me.Combo1.AddItem .Fields(1).Value
                .MoveNext
                Loop
        End If
        .Close
     Set Rs = Nothing
End With
Me.ucGradContainer1.Caption = " Secured Login..."
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdLogin_Click
End If
End Sub
