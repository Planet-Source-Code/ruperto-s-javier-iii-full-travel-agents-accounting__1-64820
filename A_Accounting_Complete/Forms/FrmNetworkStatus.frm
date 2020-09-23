VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNetworkStatus 
   Caption         =   "Verifying Network Connection"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   135
      ScaleHeight     =   1065
      ScaleWidth      =   1650
      TabIndex        =   4
      Top             =   375
      Width           =   1710
      Begin VB.Image Image1 
         Height          =   1005
         Left            =   45
         Picture         =   "FrmNetworkStatus.frx":0000
         Stretch         =   -1  'True
         Top             =   30
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   450
      Left            =   5775
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7140
      Top             =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reconnect"
      Enabled         =   0   'False
      Height          =   450
      Left            =   5775
      TabIndex        =   2
      Top             =   450
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6705
      Top             =   1725
   End
   Begin MSComCtl2.Animation anmShowAction 
      Height          =   945
      Left            =   2850
      TabIndex        =   1
      Top             =   435
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1667
      _Version        =   393216
      FullWidth       =   156
      FullHeight      =   63
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6225
      Top             =   1710
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "imgAcc Version 1.0"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1545
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Checking Network Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   7260
   End
End
Attribute VB_Name = "FrmNetworkStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim StrAvi_Filename

Private Sub Command1_Click()
Unload Me
Call Main
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
i = 0
StrAvi_Filename = App.Path & "\Avi\Connecting.avi"
anmShowAction.Open StrAvi_Filename
With anmShowAction
        .Visible = True
        .Play
End With

End Sub

Private Sub Timer1_Timer()
    If DBConnect = True Then
        
            Me.Label1.Caption = "Connected..."
            Me.Timer1.Enabled = False
            Me.Timer2.Enabled = True
    Else
            Me.Label1.Caption = "No Connection..."
            Me.Command1.Enabled = True
            With anmShowAction
                .Visible = True
                .Stop
            End With

    End If
End Sub

Private Sub Timer2_Timer()

    If DBConnect = True Then
        
           Me.Label1.Caption = "Initializing Database...Please Wait.."
           Me.Timer1.Enabled = False
           
           Me.Timer3.Enabled = True
           Me.Timer2.Enabled = False
            
    End If

End Sub

Private Sub Timer3_Timer()
           
           Me.Timer1.Enabled = False
           Me.Timer2.Enabled = False
           
             With anmShowAction
                .Visible = False
                .Stop
             End With
            Unload Me
End Sub
