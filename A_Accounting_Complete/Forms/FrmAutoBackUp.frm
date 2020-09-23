VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_AutoBackUp 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processing Backup"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
   Icon            =   "FrmAutoBackUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   8385
      TabIndex        =   4
      Top             =   0
      Width           =   8445
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Back-up Daemon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   75
         TabIndex        =   5
         Top             =   90
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   2  'Dot
      Height          =   2010
      Left            =   15
      ScaleHeight     =   1950
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   690
      Width           =   8415
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Top             =   1020
         Visible         =   0   'False
         Width           =   7920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait...  Auto Back Up in progress!!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   255
         TabIndex        =   2
         Top             =   1395
         Width           =   7935
      End
   End
End
Attribute VB_Name = "Frm_AutoBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents KulotZipMax As clsHuffman
Attribute KulotZipMax.VB_VarHelpID = -1
Dim dest As String
Dim GetAutoBackUpPath As String
Dim GetAutoBackUpName As String
Dim RsAutoBackUp As ADODB.Recordset
Dim SQL As String


Private Sub Form_Activate()
    
Set RsAutoBackUp = New ADODB.Recordset
SQL = "SELECT * FROM tbl_BackUpSettings"
RsAutoBackUp.Open SQL, cn, adOpenKeyset, adLockOptimistic
    GetAutoBackUpPath = RsAutoBackUp!AutoBackUpPath
    RsAutoBackUp.Close
    Set RsAutoBackUp = Nothing
    GetAutoBackUpName = "Data-" & Format(Now, "mm-dd-yyyy") & ".kul"
    dest = GetAutoBackUpPath & IIf(Right(GetAutoBackUpPath, 1) = "\", "", "\") & GetAutoBackUpName
    Dim OldTimer As Single
    ProgressBar1.Visible = True
    OldTimer = Timer
    Me.lblPercent.Visible = True
    Call KulotZipMax.EncodeFile(NetWorkPath & "\DbMaster.mdb", dest)
    ProgressBar1.Value = 0
    Unload Me
    Exit Sub
End Sub

Private Sub KulotZipMax_Progress(Procent As Integer)
  Label3.Caption = "Warning! Compressing Database.."
 ProgressBar1.Value = Procent
 Me.lblPercent.Caption = Procent & " % " & "Complete"
  If ProgressBar1.Value = 100 Then
    Label3.Caption = "Please Wait... System Saving Compressed File ..."
    End If
  DoEvents

End Sub

Private Sub Form_Load()
    Set KulotZipMax = New clsHuffman
End Sub

