VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_RESTORE 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Rebuilding Database Restore"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ControlBox      =   0   'False
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404040&
         Height          =   3735
         Left            =   1080
         ScaleHeight     =   3675
         ScaleWidth      =   7425
         TabIndex        =   1
         Top             =   720
         Width           =   7485
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
            Height          =   3390
            Left            =   4080
            Pattern         =   "*.SSA"
            TabIndex        =   4
            Top             =   120
            Width           =   3255
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
            Height          =   2970
            Left            =   120
            TabIndex        =   3
            Top             =   600
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
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   3735
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
         Left            =   1200
         TabIndex        =   8
         Top             =   6120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Image Image2 
         Height          =   870
         Left            =   120
         Picture         =   "frmRestore.frx":000C
         Top             =   5040
         Width           =   870
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Selection Option Mode for Auto BackUp"
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
         TabIndex        =   11
         Top             =   4680
         Width           =   8655
      End
      Begin VB.Label Label1 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   5040
         Width           =   6855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
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
         Height          =   735
         Index           =   1
         Left            =   1200
         TabIndex        =   9
         Top             =   5280
         Width           =   7335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Backed Up File and Click on  Restore Button"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   5655
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   120
         Picture         =   "frmRestore.frx":282E
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Back Up Mode for Manual Option"
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
         Width           =   8655
      End
   End
   Begin ComputerServicing.EActiveButton Cmd_Restore 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   6840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      XPBorder        =   -1  'True
      BTYPE           =   3
      TX              =   "Restore Now"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   -2147483633
      BCOLO           =   13232892
      FCOL            =   0
      FCOLO           =   128
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmRestore.frx":5050
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ComputerServicing.EActiveButton Cmd_Close 
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   6840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      XPBorder        =   -1  'True
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   -2147483633
      BCOLO           =   13232892
      FCOL            =   0
      FCOLO           =   128
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmRestore.frx":506C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FRM_RESTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub CMD_RESTORE_Click()
  If Len(Label1(1).Caption) = 0 Then
    MsgBox "Select Backup File Path and then Click Restore Button...", vbInformation, "Select Path ..."
    Exit Sub
  End If
  
  Dim OldTimer As Single
  ProgressBar1.Visible = True
  Label3.Visible = True
  
  OldTimer = Time
    Call Huffman.DecodeFile(Label1(1).Caption, App.Path & "\DataFiles\Temp.mdb")
  
'  Set f = CreateObject("Scripting.FileSystemObject")
'  f.CopyFile App.Path & "\DataFiles\Stallion.mdb", App.Path & "\DataFiles\Stallion Security Datafile.mdb", True
  
 ' f.DeleteFile App.Path & "\DataFiles\Waybill1.mdb"
 
  ProgressBar1.Value = 0
  MsgBox "New Datafile was Successfully Restored      ", vbInformation
  LogAction "Restore Back Up From : " & Label1(1).Caption
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
Label3.Visible = False
ProgressBar1.Visible = False
Set Huffman = New clsHuffman
Drive1.Drive = "c:"
End Sub

Private Sub Huffman_Progress(Procent As Integer)
  Label3.Caption = "Uncompressing Database"
  ProgressBar1.Value = Procent
  If ProgressBar1.Value = 100 Then
    Label3.Caption = "Restoring Uncompressed File ..."
    End If
  DoEvents
End Sub

