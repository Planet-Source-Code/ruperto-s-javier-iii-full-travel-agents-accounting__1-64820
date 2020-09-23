VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim frmBusy As New frmBusy 'lets asycronus processing

Public Sub BusyStatus(ByVal Status As Integer, strMessage As String)
'byval means sending actual contents of memory location that argument has
'gives copy of item in memory location
    
    On Error Resume Next 'can receive neg values so we continue
    
    With frmBusy
        .Visible = True
        .Cls
        .ucGradContainer1.Caption = strMessage          'change message
        .shpTop.Width = .Width * (Status / 100) 'receive (0-100) percent length of form
        .shpBottom.Width = .shpTop.Width - 75   'bottom will be less to cast shadow
    End With

End Sub

Public Sub BusyExit()
    With frmBusy
        .Visible = True
        .shpTop.Width = frmBusy.Width
        .shpBottom.Width = .shpTop.Width
        .ucGradContainer1.Caption = "Complete..."
    End With
    Sleep 250
    Unload frmBusy
    
End Sub
