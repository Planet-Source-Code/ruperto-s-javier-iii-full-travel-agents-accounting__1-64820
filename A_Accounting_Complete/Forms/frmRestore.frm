VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmRestore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Restore"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   375
      Left            =   5250
      TabIndex        =   11
      Top             =   6435
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
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
      MICON           =   "frmRestore.frx":000C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdRestore 
      Height          =   375
      Left            =   2925
      TabIndex        =   10
      Top             =   6435
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Restore"
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
      MICON           =   "frmRestore.frx":0028
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   5610
      Left            =   0
      ScaleHeight     =   5550
      ScaleWidth      =   7590
      TabIndex        =   0
      Top             =   0
      Width           =   7650
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00800000&
         Height          =   3765
         Left            =   60
         ScaleHeight     =   3705
         ScaleWidth      =   7485
         TabIndex        =   1
         Top             =   570
         Width           =   7545
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3150
            Left            =   3765
            Pattern         =   "*.kul"
            TabIndex        =   4
            Top             =   30
            Width           =   3690
         End
         Begin VB.DirListBox Dir1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   15
            TabIndex        =   3
            Top             =   435
            Width           =   3735
         End
         Begin VB.DriveListBox Drive1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   30
            TabIndex        =   2
            Top             =   45
            Width           =   3705
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Back-Up Archive List"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4365
            TabIndex        =   13
            Top             =   3225
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Archive Explorer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   795
            TabIndex        =   12
            Top             =   3240
            Width           =   1845
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
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   4080
            Width           =   3015
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   5130
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Back Up File Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   4320
         Width           =   7530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   540
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   4590
         Width           =   7545
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Restore Previous Save Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7680
      End
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
      Left            =   90
      TabIndex        =   15
      Top             =   5625
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      Height          =   285
      Left            =   60
      TabIndex        =   14
      Top             =   6060
      Visible         =   0   'False
      Width           =   7560
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents KulotZipMax As clsHuffman
Attribute KulotZipMax.VB_VarHelpID = -1

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRestore_Click()
  If Len(Label1(1).Caption) = 0 Then
    MsgBox "Select Backup File Path and then Click Restore Button...", vbInformation, "Select Path ..."
    Exit Sub
  End If
  
  Dim OldTimer As Single
  ProgressBar1.Visible = True
  Label3.Visible = True
  Me.Label2.Visible = True
  Me.lblPercent.Visible = True
  OldTimer = Time
  Call KulotZipMax.DecodeFile(Label1(1).Caption, App.Path & "\Data\Temp.mdb")
  
  ProgressBar1.Value = 0
  MsgBox "New Datafile was Successfully Restored      ", vbInformation
  'LogAction "Restore Back Up From : " & Label1(1).Caption
  MsgBox "System will restart to initialize restored data"
  End
  Unload Me
  Exit Sub
End Sub

Private Sub Dir1_Change()
Label1(1).Caption = ""
On Error GoTo A1:
    File1.Path = Dir1.Path
    Exit Sub
A1:
    MsgBox "Folder Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Drive1_Change()
Label1(1).Caption = ""
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"

End Sub

Private Sub File1_Click()
    Label1(1).Caption = File1.Path & IIf(Right(File1.Path, 1) = "\", "", "\") & File1.Filename
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If ProgressBar1.Value = 0 Or ProgressBar1.Value Then
        Unload Me
    End If
End If

End Sub

Private Sub Form_Load()
KeyPreview = True

Me.Top = 0
Me.Left = 0
Label1(1).Caption = ""
Label2.Visible = False
ProgressBar1.Visible = False
Me.lblPercent.Visible = False
Set KulotZipMax = New clsHuffman
Drive1.Drive = "c:"
End Sub

Private Sub KulotZipMax_Progress(Procent As Integer)
  Label2.Caption = "Uncompressing Database"
  Me.lblPercent.Caption = Procent & " % " & "Complete"
  ProgressBar1.Value = Procent
  If ProgressBar1.Value = 100 Then
    Label2.Caption = "Restoring Uncompressed File Complete..."
    End If
  DoEvents
End Sub

