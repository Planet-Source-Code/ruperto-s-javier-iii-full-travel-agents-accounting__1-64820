VERSION 5.00
Begin VB.Form frmSelect 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   390
      Left            =   1200
      TabIndex        =   3
      Top             =   1950
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3570
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "One Stub Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   810
         TabIndex        =   2
         Top             =   870
         Width           =   2190
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Individual Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   795
         TabIndex        =   1
         Top             =   285
         Width           =   2190
      End
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Me.Option1 Then
'MsgBox Mid(frmStatement.GetLastNumber, 1, ReturnFirst(frmStatement.GetLastNumber))

If ReturnFirst(frmStatement.GetLastNumber) = 0 Then
    'frmStatement.txtNo = AutoIncrement(frmStatement.GetLastNumber) & "-" & kulotRead(App.Path & "\Settings.txt")
    frmStatement.txtNo = frmStatement.GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
Else
    'frmStatement.txtNo = AutoIncrement(Mid(frmStatement.GetLastNumber, 1, ReturnFirst(frmStatement.GetLastNumber))) & kulotRead(App.Path & "\Settings.txt")
    frmStatement.txtNo = frmStatement.GetLastNumber & "-" & kulotRead(App.Path & "\Settings.txt")
End If
    Unload Me
    Exit Sub
End If
If Me.Option2 Then
    frmAsk.Show 1
    Unload Me
    
End If
End Sub

