VERSION 5.00
Begin VB.Form frmBusy 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   DrawMode        =   2  'Blackness
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Accounting.ucGradContainer ucGradContainer1 
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1799
      IconSize        =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderIcon      =   "frmBusy.frx":0000
      Begin VB.Shape shpTop 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   -45
         Top             =   510
         Width           =   45
      End
      Begin VB.Shape shpBottom 
         BorderColor     =   &H8000000B&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   -30
         Top             =   555
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

