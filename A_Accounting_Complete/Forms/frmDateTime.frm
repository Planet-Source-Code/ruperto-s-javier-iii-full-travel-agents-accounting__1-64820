VERSION 5.00
Begin VB.Form frmDateTime 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDateTime.frx":0000
   ScaleHeight     =   3630
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "F3 - &Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4560
      Picture         =   "frmDateTime.frx":2B2B
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F6 - &Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3000
      Picture         =   "frmDateTime.frx":34C5
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please take time to check the system date and time before starting any transaction."
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "System Date:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "System Date:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   1605
      Left            =   4080
      Picture         =   "frmDateTime.frx":3E5F
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   3630
      Left            =   0
      Top             =   0
      Width           =   6180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "System Time:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "System Date:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
        Load frmLogin
        frmLogin.Show 1
    Case 1
        Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus

End Select
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Unload Me
End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
lbl(2) = Format(Date, "mmmm dd, yyyy")
lbl(3) = Format(Time, "hh:mm:ss")

End Sub
