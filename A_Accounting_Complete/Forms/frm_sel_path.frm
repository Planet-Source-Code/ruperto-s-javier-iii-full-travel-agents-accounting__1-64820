VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frm_sel_path 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Mode"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   9705
   Icon            =   "frm_sel_path.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400040&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9675
      TabIndex        =   6
      Top             =   0
      Width           =   9735
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frm_sel_path.frx":08CA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " Application Path"
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
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Path Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   90
      TabIndex        =   1
      Top             =   885
      Width           =   9495
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Path ..."
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
         Left            =   4560
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1935
         Index           =   1
         Left            =   4560
         TabIndex        =   4
         Top             =   960
         Width           =   4770
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   90
      TabIndex        =   0
      Top             =   4185
      Width           =   9495
      Begin LVbuttons.LaVolpeButton cmdClose 
         Height          =   555
         Left            =   7050
         TabIndex        =   9
         Top             =   255
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   979
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
         MICON           =   "frm_sel_path.frx":1194
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
      Begin LVbuttons.LaVolpeButton cmdOk 
         Height          =   555
         Left            =   4785
         TabIndex        =   8
         Top             =   255
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   979
         BTYPE           =   2
         TX              =   "Ok"
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
         MICON           =   "frm_sel_path.frx":11B0
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
End
Attribute VB_Name = "frm_sel_path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Dim MyForm As Form
Dim ObjectName As String

Private Sub Cmd_Close_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
MyForm.Controls(ObjectName).Text = Label1(1).Caption
Unload Me
End Sub

Private Sub Dir1_Change()
Label1(1).Caption = Dir1.Path
End Sub



Private Sub Drive1_Change()
On Error GoTo A1:
Dir1.Path = Drive1.Drive
Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Form_Load()
Drive1.Drive = "c:"
Label1(1).Caption = Dir1.Path
End Sub

Private Sub Cmd_Ok_Click()

End Sub

Public Sub SetSource(cForm As Form, cObjectName As String)
    Set MyForm = cForm
    ObjectName = cObjectName
End Sub
