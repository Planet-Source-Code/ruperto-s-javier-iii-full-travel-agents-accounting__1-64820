VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLoading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3495
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Accounting.ucGradContainer ucGradContainer1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _extentx        =   16325
      _extenty        =   6165
      headericon      =   "frmLoading.frx":0000
      captionfont     =   "frmLoading.frx":6864
      iconsize        =   0
      backcolor1      =   16577775
      caption         =   "ELS Travel and Tours Ledesma St., Iloilo City"
      captionalignment=   2
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8190
         Top             =   1140
      End
      Begin MSComCtl2.Animation anmShowAction 
         Height          =   765
         Left            =   4095
         TabIndex        =   1
         Top             =   1950
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1349
         _Version        =   393216
         BackColor       =   16577775
         FullWidth       =   113
         FullHeight      =   51
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Data... Please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2895
         TabIndex        =   3
         Top             =   1230
         Width           =   4230
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) 2005 imaginaSoft All Rights Reserved"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   105
         TabIndex        =   2
         Top             =   3210
         Width           =   8595
      End
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Private Sub Form_Load()
Dim StrAvi_Filename As String
Dim Rs As Recordset
StrAvi_Filename = App.Path & "\Avi\SEARCH.AVI"
anmShowAction.Open StrAvi_Filename
With anmShowAction
        .Visible = True
        .Play
End With

Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM tbl_Users"
Rs.Open SQL, cn, adOpenKeyset, adLockOptimistic
If Rs.RecordCount > 0 Then
    Rs.Close
End If
Set Rs = Nothing
ctr = 0
Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ctr = ctr + 1
DoEvents
If ctr >= 10 Then
With anmShowAction
        .Visible = False
        .Stop
End With
    Unload Me
End If
End Sub
