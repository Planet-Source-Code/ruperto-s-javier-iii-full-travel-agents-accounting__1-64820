Attribute VB_Name = "modINI"
'//---------------------------------------------------------------------------------------
'//--Module    : modINI
'//--DateTime  : 11.02.2005
'//--Author    : Ruperto S. Javier III a.k.a [boykulot]
'//--Purpose   : INI Files read write
'//---------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'

Public Function AppPath() As String
    AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\")
End Function

Public Sub SaveINI(INIFile As String, INIHead As String, INIKey As String, INIVal As String)
  
    Dim INIFileName As String
    Dim sRet        As String
  
    INIFileName = AppPath & INIFile
    sRet = WritePrivateProfileString(INIHead, INIKey, INIVal, INIFileName)
End Sub

Public Function GetINI(INIFile As String, INIHead As String, INIKey As String, INIDefault As String) As String

    Dim INIFileName As String
    Dim Temp        As String * 260
    Dim sRet        As String
    
    INIFileName = AppPath & INIFile
    sRet = GetPrivateProfileString(INIHead, INIKey, INIDefault, Temp, Len(Temp), INIFileName)
    GetINI = Trim$(Temp)
    GetINI = Left$(GetINI, Len(GetINI) - 1)
End Function

