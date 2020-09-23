VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmBackUpData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Back Up"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10860
   Icon            =   "frmBackUp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1620
      Left            =   45
      TabIndex        =   16
      Top             =   5130
      Width           =   10740
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   60
         TabIndex        =   17
         Top             =   180
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   609
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
         Left            =   1590
         TabIndex        =   19
         Top             =   570
         Visible         =   0   'False
         Width           =   7500
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1140
         Width           =   10530
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10830
      TabIndex        =   10
      Top             =   -15
      Width           =   10890
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Back-up Database"
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
         Left            =   390
         TabIndex        =   11
         Top             =   225
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   45
      TabIndex        =   9
      Top             =   6825
      Width           =   10740
      Begin LVbuttons.LaVolpeButton cmdClose 
         Height          =   660
         Left            =   7620
         TabIndex        =   13
         Top             =   225
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1164
         BTYPE           =   2
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
         MICON           =   "frmBackUp.frx":1CFA
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
      Begin LVbuttons.LaVolpeButton cmdBackItUp 
         Height          =   615
         Left            =   345
         TabIndex        =   14
         Top             =   300
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   "Back - Up Now"
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
         MICON           =   "frmBackUp.frx":1D16
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Auto BackUp Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4125
      Left            =   5625
      TabIndex        =   5
      Top             =   960
      Width           =   5160
      Begin VB.CheckBox Check2 
         Caption         =   "Activate Timed Auto BackUp "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   870
         Width           =   4215
      End
      Begin VB.Frame Frame5 
         Height          =   1350
         Left            =   135
         TabIndex        =   20
         Top             =   1230
         Width           =   4920
         Begin VB.TextBox txtHour 
            Height          =   330
            Left            =   1380
            TabIndex        =   22
            Text            =   "0"
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "TIME :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   420
            TabIndex        =   21
            Top             =   390
            Width           =   765
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Activate Auto BackUp On System Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   435
         Width           =   4215
      End
      Begin VB.TextBox txtAutoBackUpPath 
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
         Left            =   135
         TabIndex        =   6
         Top             =   2925
         Width           =   4890
      End
      Begin LVbuttons.LaVolpeButton cmdSettings 
         Height          =   645
         Left            =   2040
         TabIndex        =   15
         Top             =   3360
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1138
         BTYPE           =   2
         TX              =   "Update Settings"
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
         MICON           =   "frmBackUp.frx":1D32
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
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Path to Auto Back Up..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   2640
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Manual Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4125
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   5520
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   4980
         TabIndex        =   12
         Top             =   1335
         Width           =   375
      End
      Begin VB.TextBox Text2 
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   4545
      End
      Begin VB.TextBox txtAutoBackUpName 
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
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   4545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Path Folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Database FileName (Filename BackUp Only)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmBackUpData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents KulotZipMax As clsHuffman
Attribute KulotZipMax.VB_VarHelpID = -1
Dim SQL As String

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.Check1.Value = 1 Then
        GetSetPath
     '   Me.cmdAutoBackUpPath.Enabled = True
'    Else
'        Me.txtAutoBackUpPath.Text = ""
    End If
End Sub
Private Sub GetSetPath()
    Dim GetAutoBackUpPath As String, GetAutoBackUpName As String
    GetAutoBackUpPath = RsBackUpSettings!AutoBackUpPath
    Me.txtAutoBackUpPath = Trim(GetAutoBackUpPath)
End Sub

Private Sub cmdBackItUp_Click()
If Len(Me.txtAutoBackUpName.Text) = 0 Then
    MsgBox "Cannot Perform Back Up without Back Up FileName", vbCritical
    Exit Sub
End If
If Len(Text2.Text) = 0 Then
    MsgBox "Select Backup File Path and then Click Create Backup Button...", vbInformation, "Select Backup Path ..."
    Exit Sub
End If

  Label3.Visible = True
  
  Dim dest As String
  dest = Text2.Text & IIf(Right(Text2.Text, 1) = "\", "", "\") & Me.txtAutoBackUpName.Text & ".kul"
  Dim OldTimer As Single
  ProgressBar1.Visible = True
  Me.lblPercent.Visible = True
  OldTimer = Timer
  
  Call KulotZipMax.EncodeFile(NetWorkPath & "\DbMaster.mdb", dest)
  ProgressBar1.Value = 0
  MsgBox "   Back Up was successfully created       ", vbInformation
 'WriteToLog "Create Back Up To : " & dest
  Unload Me
  Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSettings_Click()
        If Me.Check1.Value = 1 Then
            If Len(Me.txtAutoBackUpPath.Text) = 0 Then
                MsgBox "   Cannot Update without Auto Back Up Path       ", vbCritical
                GoTo ExitSub
            End If
        End If
        
        If Me.Check2.Value = 1 Then
            If Len(Me.txtHour) = 0 Then
                MsgBox "   Cannot Update without Auto Back Up Time       ", vbCritical
                GoTo ExitSub
            End If
            
            
            
        End If
        UpdateAutoBackUpSettings
        Me.cmdSettings.Caption = "Modify Settings"
        MsgBox "   Settings was successfully Update       ", vbInformation
ExitSub:

End Sub

Private Sub Command1_Click()
    frm_sel_path.SetSource Me, "Text2"
    frm_sel_path.Show vbModal
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If ProgressBar1.Value = 0 Or ProgressBar1.Value Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()


    GetSetPath
    Me.txtAutoBackUpName.Text = "Data-" & Format(Now, "mm-dd-yyyy")
    
    If RsBackUpSettings!BackUpOnClose = True Then
        Me.Check1.Value = 1
    Else
        Me.Check1.Value = 0
    End If
    
    If RsBackUpSettings!BackUpOnTime = True Then
        Me.Check2.Value = 1
    Else
        Me.Check2.Value = 0
    End If
    

        Me.txtHour = Format(RsBackUpSettings!cTime, "hh:mm AMPM")

    
KeyPreview = True
Set KulotZipMax = New clsHuffman
Label3.Visible = False
ProgressBar1.Visible = False

End Sub

Private Sub UpdateAutoBackUpSettings()
                                                                                                                    'WriteStringValue MyRegKey, "AutoBackUpPath", IIf(Len(Me.txtAutoBackUpPath.Text) = 0, " ", Me.txtAutoBackUpPath.Text)
        
        RsBackUpSettings!AutoBackUpPath = txtAutoBackUpPath
        RsBackUpSettings!BackUpOnClose = Me.Check1.Value
        RsBackUpSettings!BackUpOnTime = Me.Check2.Value
        RsBackUpSettings!cTime = Format(Me.txtHour, "hh:mm AMPM")
        RsBackUpSettings.Update
        
End Sub

Private Sub Image1_Click()

End Sub

Private Sub KulotZipMax_Progress(Procent As Integer)
 Label3.Caption = "Compressing Database... Please Wait.."
 ProgressBar1.Value = Procent
 Me.lblPercent.Caption = Procent & " % " & "Complete"
  If ProgressBar1.Value = 100 Then
    Label3.Caption = "Saving Compressed File ..."
    End If
  DoEvents
End Sub

Private Sub txtAutoBackUpName_GotFocus()
    HighlightText Me.txtAutoBackUpName
End Sub

Private Sub txtAutoBackUpName_KeyPress(KeyAscii As Integer)
    Alphabets KeyAscii
End Sub


Public Sub HighlightText(myTxt As Object)
    On Error GoTo ExitSub
    With myTxt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
ExitSub:
End Sub

Public Sub Alphabets(KeyAscii As Integer)
' this is for Alphabet validation
    If Not (KeyAscii > 96 And KeyAscii < 123) Then
        If Not (KeyAscii > 64 And KeyAscii < 91) Then
            If Not KeyAscii = Asc(" ") Then
                If Not KeyAscii = vbKeyBack Then
                    KeyAscii = 0
                End If
            End If
        End If
    End If
End Sub
