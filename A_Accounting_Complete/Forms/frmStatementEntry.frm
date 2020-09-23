VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatementEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entry"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   495
      Left            =   90
      TabIndex        =   5
      Top             =   4080
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9450
      TabIndex        =   14
      Top             =   4095
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   3840
      Left            =   90
      TabIndex        =   6
      Top             =   195
      Width           =   10890
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shipping / Airline / Ticket Info"
         Enabled         =   0   'False
         Height          =   2160
         Left            =   6030
         TabIndex        =   15
         Top             =   1500
         Width           =   4665
         Begin VB.TextBox txtPRice 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1755
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Height          =   345
            Left            =   1755
            TabIndex        =   19
            Top             =   885
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Height          =   345
            Left            =   1755
            TabIndex        =   18
            Top             =   360
            Width           =   2670
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Price :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   345
            TabIndex        =   20
            Top             =   1620
            Width           =   1485
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Ticket Type :"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   315
            TabIndex        =   17
            Top             =   915
            Width           =   1485
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Shipping / Airline :"
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   300
            TabIndex        =   16
            Top             =   405
            Width           =   1605
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Route Info"
         Height          =   2160
         Left            =   180
         TabIndex        =   12
         Top             =   1500
         Width           =   5760
         Begin VB.TextBox txtRoute 
            Height          =   375
            Left            =   1350
            TabIndex        =   4
            Top             =   360
            Width           =   4275
         End
         Begin VB.Label Label5 
            Caption         =   "Route"
            Height          =   270
            Left            =   555
            TabIndex        =   13
            Top             =   375
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Departure Info"
         Height          =   1110
         Left            =   5955
         TabIndex        =   9
         Top             =   285
         Width           =   4785
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   1620
            TabIndex        =   2
            Top             =   225
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   609
            _Version        =   393216
            Format          =   19529729
            CurrentDate     =   38474
         End
         Begin VB.TextBox txtDepartureTime 
            Height          =   345
            Left            =   1620
            TabIndex        =   3
            Text            =   "12:00:00 AM"
            Top             =   630
            Width           =   3000
         End
         Begin VB.Label Label4 
            Caption         =   "Departure Time :"
            Height          =   270
            Left            =   255
            TabIndex        =   11
            Top             =   645
            Width           =   1470
         End
         Begin VB.Label Label3 
            Caption         =   "Departure Date :"
            Height          =   270
            Left            =   255
            TabIndex        =   10
            Top             =   270
            Width           =   1470
         End
      End
      Begin VB.TextBox txtTicketNo 
         Height          =   345
         Left            =   1530
         TabIndex        =   0
         Top             =   450
         Width           =   4290
      End
      Begin VB.TextBox txtPassenger 
         Height          =   345
         Left            =   1530
         TabIndex        =   1
         Top             =   930
         Width           =   4290
      End
      Begin VB.Label Label2 
         Caption         =   "Ticket No. :"
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   465
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Passenger Name :"
         Height          =   270
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmStatementEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
