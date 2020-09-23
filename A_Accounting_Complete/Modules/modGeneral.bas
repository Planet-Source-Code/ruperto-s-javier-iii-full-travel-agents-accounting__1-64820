Attribute VB_Name = "modGeneral"
'//---------------------------------------------------------------------------------------
'//--Module    : modGeneral
'//--DateTime  : 11.02.2005
'//--Author    : Ruperto S. Javier III a.k.a [boykulot]
'//--Purpose   : INI Files read write
'//---------------------------------------------------------------------------------------
'API Subs
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Sub Sleep Lib "kernel32" (ByVal MilliSeconds As Long)
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As meminfo_status)

'API Functions
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, Y, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function RegControl Lib "c:\windows\system32\kulotmenu.OCX" Alias "DllRegisterServer" () As Long
Public Declare Function UnRegControl Lib "c:\windows\system32\kulotmenu.OCX" Alias "DllUnregisterServer" () As Long
  


Public Const S_OK = &H0
Public Const MF_BYPOSITION = &H400&
Private Const MAX_PATH = 260

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SW_MINIMIZE = 6
Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Public Enum FilePartTypes
    FileExtOnly
    FileNameOnly
    FileNameAndExt
    FilePathOnly
End Enum


Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Public Type meminfo_status
    dwlength                    As Long
    dwmemoryload                As Long
    dwtotalphy                  As Long
    dwavaiphy                   As Long
    dwtotalpagefile             As Long
    dwavaipagefile              As Long
    dwtotalvirtual              As Long
    dwavailabelvirtual          As Long
End Type

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type



Public meminfo As meminfo_status

Public Function Return_AccNo(Param) As String
Dim Rst As New ADODB.Recordset
Dim SQL As String
SQL = "SELECT DISTINCT  [Account Number],[Account Name] FROM tbl_AccountsSetting WHERE [Account Name]='" & Param & "' ORDER by [Account Number] ASC"
With Rst
        .Open SQL, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Return_AccNo = .Fields("Account Number").Value
        End If
End With
End Function

Public Sub RegisterControlSub()
On Error GoTo Err_Registration_Failed
   If RegControl = S_OK Then
      MsgBox "Registered Successfully"
   Else
      MsgBox "Not Registered"
   End If
   Exit Sub
Err_Registration_Failed:
   MsgBox "Error: " & Err.Number & " " & Err.Description
End Sub

Sub UnRegisterControlSub()
   On Error GoTo Err_Unregistration_Failed
   If UnRegControl = S_OK Then
      MsgBox "Unregistered Successfully"
   Else
      MsgBox "Not Unregistered"
   End If
   Exit Sub
Err_Unregistration_Failed:
   MsgBox "Error: " & Err.Number & " " & Err.Description
End Sub


Public Sub RemoveButtonX(frm As Form)
On Error GoTo FailSafe_Error
    Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(frm.hWnd, 0)
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
Exit Sub
FailSafe_Error:
End Sub



Public Function IsNumber(KeyAscii As Integer) As Boolean
On Error GoTo FailSafe_Error
10          If InStr(1, "1234567890.", Chr$(KeyAscii)) Or KeyAscii = 8 Then
20              IsNumber = True
30          Else
40              IsNumber = False
50          End If
Exit Function
FailSafe_Error:
End Function

Public Function CheckNull(obj) As Boolean
On Error GoTo FailSafe_Error
10          If Len(obj) = 0 Then
20              CheckNull = True
30            Else
40              CheckNull = False
50          End If
Exit Function
FailSafe_Error:
End Function

Sub TrapEnter(KeyAscii As Integer)
On Error GoTo FailSafe_Error
10          If KeyAscii = 13 Then
20               SendKeys "{Tab}"
30          End If
Exit Sub
FailSafe_Error:
End Sub


Public Function AutoIncrement(ByVal Param As String)
On Error GoTo FailSafe_Error
    Dim strlen As Integer
    Dim tmpStr() As String
    strlen = Len(Param) ' number of characters
    ReDim tmpStr(strlen)

    For L = 1 To UBound(tmpStr) ' parse individual characters
        tmpStr(L) = Mid(Param, L, 1)
    Next L
    
    For nxtchar = 1 To UBound(tmpStr) ' cyle through characters increment ascii value
        valchar = (UBound(tmpStr)) - (nxtchar - 1)
        If Asc(tmpStr(valchar)) >= 65 And Asc(tmpStr(valchar)) <= 90 Or _
        Asc(tmpStr(valchar)) >= 97 And Asc(tmpStr(valchar)) <= 122 Then ' upper and lower alpha characters


        If Asc(tmpStr(valchar)) = 90 Or Asc(tmpStr(valchar)) = 122 Then


            If Asc(tmpStr(valchar)) = 90 Then


                If valchar = 1 Then ' fisrt char at the End of ascii list
                    tmpStr(valchar) = "AA"
                Else
                    tmpStr(valchar) = "A"
                End If

            Else


                If valchar = 1 Then ' fisrt char at the End of ascii list
                    tmpStr(valchar) = "aa"
                Else
                    tmpStr(valchar) = "a"
                End If

            End If

        Else
            tmpStr(valchar) = Chr(Asc(tmpStr(valchar)) + 1) ' increment ascii by one
            GoTo noneedto:
        End If

    ElseIf Asc(tmpStr(valchar)) > 47 And Asc(tmpStr(valchar)) < 58 Then 'numeric values


        If Asc(tmpStr(valchar)) = 57 Then


            If valchar = 1 Then ' fisrt char at the End of ascii list
                tmpStr(valchar) = "10"
            Else
                tmpStr(valchar) = "0"
            End If

        Else
            tmpStr(valchar) = Chr(Asc(tmpStr(valchar)) + 1) ' increment ascii by one
            GoTo noneedto:
        End If

    End If

Next nxtchar

noneedto: 'once a char is increment and is Not carried over no need to increment all chars


For mke = LBound(tmpStr) To UBound(tmpStr) ' make text
    AutoIncrement = Trim$(AutoIncrement) & tmpStr(mke)
Next mke
Exit Function
FailSafe_Error:
End Function

Public Function ReturnFirst(Param) As Integer
Dim i As Integer
Dim ctr As Integer

ctr = 0
For i = 1 To Len(Param) Step 1
    If Mid(Param, i, 1) = "-" Then
    ctr = ctr + 1
      If ctr = 2 Then
        ReturnFirst = i
      End If
    End If
Next i

End Function


Public Function Return_1stDash(Param) As Integer
Dim i As Integer
For i = 1 To Len(Param) Step 1
    If Mid(Param, i, 1) = "-" Then
        Return_1stDash = i
        Exit For
    End If
Next i

End Function


Public Function returnMon() As String
Dim Tmp
If Len(Month(Now)) <= 1 Then
Tmp = "0" & Month(Now)
Else
Tmp = Month(Now)
End If
returnMon = Tmp
End Function

Public Function returnDay() As String
Dim Tmp

If Len(Day(Now)) <= 1 Then
Tmp = "0" & Day(Now)
Else
Tmp = Day(Now)
End If
returnDay = Tmp
End Function

Public Function Encrypt(ByVal MyStr As String) As String
    Dim i As Integer
    Encrypt = ""
    For i = 1 To Len(MyStr)
        Encrypt = Encrypt & Chr(Asc(Mid(MyStr, i, 1)) + 1)
    Next i
End Function

Public Function Decrypt(ByVal MyStr As String) As String
    Dim i As Integer
    Decrypt = ""
    For i = 1 To Len(MyStr)
        Decrypt = Decrypt & Chr(Asc(Mid(MyStr, i, 1)) - 1)
    Next i
End Function

'Procedure used to check all loaded form in memory
Function IsLoaded(ByVal frm As Form) As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = frm.Name Then
            IsLoaded = True
            Exit Function
        End If
    Next
    IsLoaded = False
End Function

'Procedure used to highlight text when focus
Public Sub kulotHL(ByRef theText As Object)
    On Error Resume Next
    With theText
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

'Procedure used to clear the text content
Public Sub kulotClrText(ByRef sForm As Form)
    Dim theControl As Control
    For Each theControl In sForm.Controls
        If (TypeOf theControl Is TextBox) Then theControl = vbNullString
    Next theControl
    Set theControl = Nothing
End Sub

'Function used to read from txt files
Public Function kulotRead(FN As String) As String
    Dim i As Integer
    i = FreeFile
    On Error GoTo FailSafe_Error
    Open FN For Input As #i
    kulotRead = Input(LOF(i), i)
    Close #i
    Exit Function
FailSafe_Error:
    kulotRead = ""
End Function

'Procedure used to write to txt files
Public Sub kulotWrite(FN As String, Contents As String)
    Dim i As Integer
    i = FreeFile
    On Error GoTo FailSafe_Error
    Open FN For Output As #i
    Print #i, Contents
    Close #i
    Exit Sub
FailSafe_Error:
    MsgBox "Error writing to file :" & FN & " " & Err.Description
End Sub


Public Function GetDecimalSeparator() As String
    GetDecimalSeparator = Mid(3 / 2, 2, 1)
End Function


Public Function RoundToDecimalPosition(ByVal dblValue As Double, lngPosition As Long)
    Dim lngDecimalPosition As Long
    Dim strDecimalSeparator As String
    'The following is programming language i
    '     ndependent
    strDecimalSeparator = GetDecimalSeparator
    lngDecimalPosition = InStr(1, dblValue, strDecimalSeparator)


    If Len(dblValue) - lngDecimalPosition > lngPosition Then
        'Adding or substracting 0.5 allows for r
        '     ounding instead of truncating
        dblValue = (10 ^ lngPosition) * dblValue + IIf(dblValue < 0, -0.5, 0.5)
        lngDecimalPosition = InStr(1, dblValue, strDecimalSeparator)
        If lngDecimalPosition Then dblValue = Left(dblValue, lngDecimalPosition - 1)
        'if there is not decimal then there is n
        '     o need to strip it
        dblValue = dblValue / (10 ^ lngPosition)
    End If
    RoundToDecimalPosition = dblValue
End Function


Function r(ByVal x As Double, ByVal n As Integer) As Double

r = CDbl(Int(x * 10 ^ n) + 0.5 / 10 ^ n)
End Function


Public Function roundDown(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
    roundDown = CDbl(Left(CStr(dblValue), myDec))
Else
    roundDown = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Down"
End Function

Public Function roundUp(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
    roundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
Else
    roundUp = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Up"
End Function

Public Sub ShortCut()

    On Error Resume Next

    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")

    Dim MyShortcut, _
        MyDesktop, _
        DesktopPath

    DesktopPath = WSHShell.specialfolders("AllUsersDesktop")

    Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\" & ProjectName & ".lnk")

    With MyShortcut
        .targetpath = WSHShell.ExpandEnvironmentStrings(PathCheck(App.Path)) & ProjectName & ".exe"
        .WorkingDirectory = WSHShell.ExpandEnvironmentStrings(PathCheck(App.Path))
        .WindowStyle = 4
        .IconLocation = WSHShell.ExpandEnvironmentStrings(PathCheck(App.Path) & ProjectName & ".exe") & ", 0"
        .Save
    End With

    Err.Clear

End Sub


'+-------------------+
'|DATE TIME Functions|
'+-------------------+
Public Function BusinessDays(ByVal Date1 As Variant, Optional ByVal Date2 As Variant, Optional WeekEndsOnly As Boolean = False) As Long

    If Not IsDate(Date1) Then Exit Function

    If Not IsDate(Date2) Then Date2 = Date

    Date1 = CDate(Date1)
    Date2 = CDate(Date2)

    If Date1 > Date2 Then
        'Swap values
        Dim lTemp As Date
        lTemp = Date2
        Date2 = Date1
        Date1 = lTemp
    End If

    BusinessDays = 0

    If WeekEndsOnly Then    'ONLY count the Days on the WeekEnd(s)
        While Date1 <= Date2
            If IsWeekend(Date1) Then BusinessDays = BusinessDays + 1
            Date1 = Date1 + 1
        Wend
    Else                    'ONLY count the Days during the Week
        While Date1 <= Date2
            If Not IsWeekend(Date1) Then BusinessDays = BusinessDays + 1
            Date1 = Date1 + 1
        Wend
    End If

End Function

Function BusinessDaysAdd(ByVal SomeDate As Date, ByVal Days As Long, Optional ByVal SaturdayIsHoliday As Boolean = True) As Date

    Do While Days
        SomeDate = SomeDate + Sgn(Days)   ' increment or decrement the date
        ' check that it is a week day
        If Weekday(SomeDate) <> vbSunday And (Weekday(SomeDate) <> vbSaturday Or Not SaturdayIsHoliday) Then
            ' days becomes closer to zero
            Days = Days - Sgn(Days)
        End If
    Loop

    BusinessDaysAdd = SomeDate

End Function



Public Function DaysInYear(ByVal SomeValue As Variant) As Integer
    If IsDate(SomeValue) Or IsNumeric(SomeValue) Then DaysInYear = IIf(IsLeapYear(SomeValue), 366, 365)
End Function

Public Function IsLeapYear(ByVal SomeValue As Variant) As Boolean

    On Error GoTo LocalError

    Dim intYear As Integer

    'The 3 Golden rules are:
    '1. True if it is divisible by 4
    '2. False if it is divisible by 100
    '3. TRUE if it is divisble by 400
    If IsDate(SomeValue) Then intYear = Year(SomeValue) Else intYear = CInt(SomeValue)

    If TypeName(intYear) = "Integer" Then
        'Using DateSerial Function
        IsLeapYear = Day(DateSerial(intYear, 3, 0)) = 29
        'IsLeapYear = Day(DateSerial(intYear, 2, 29)) = 29
        'Using Calculations
        'IsLeapYear = ((intYear Mod 4 = 0) And (intYear Mod 100 <> 0) Or (intYear Mod 400 = 0))
    End If
Exit Function

LocalError:
End Function

Public Function DayPart(Optional vTime As Variant = "", Optional Greeting As String) As String

    If vTime = "" Then vTime = Time

    If IsDate(vTime) Then vTime = FormatDateTime(vTime, vbShortTime)

    If (vTime >= #12:00:00 AM#) And (vTime < #12:00:00 PM#) Then
        DayPart = "morning"
    ElseIf (vTime > #12:00:00 AM#) And (vTime < #5:00:00 PM#) Then
        DayPart = "afternoon"
    Else
        DayPart = "evening"
    End If

    If Len(Greeting) > 0 Then DayPart = TrimALL(Greeting & " " & DayPart)

End Function

Public Function EndOfMonth(SomeDate As Variant) As Date

    If IsDate(SomeDate) Then
        EndOfMonth = DateAdd("m", 1, SomeDate)
        EndOfMonth = DateSerial(Year(EndOfMonth), Month(EndOfMonth), 1)
        EndOfMonth = DateAdd("d", -1, EndOfMonth)
    End If

End Function

Function EndOfWeek(ByVal SomeDate As Date) As Date

    If IsDate(SomeDate) Then
        EndOfWeek = FormatDateTime(SomeDate - Weekday(SomeDate) + 7, vbGeneralDate)
    End If

End Function

Public Function IsWeekend(ByVal SomeDate As Variant) As Boolean

    If IsDate(SomeDate) Then If (Weekday(SomeDate) = 1) Or (Weekday(SomeDate) = 7) Then IsWeekend = True

End Function

Public Function NextDate(ByVal d As Date, Optional ByVal WhatDay As VbDayOfWeek = vbSaturday, Optional GetNext As Boolean = True) As Date

    NextDate = (((d - WhatDay + GetNext) \ 7) - GetNext) * 7 + WhatDay

End Function

Public Function WeekNo(Optional SomeDate As Variant) As Integer

    WeekNo = DatePart("ww", IIf(IsDate(SomeDate), SomeDate, Date))

End Function
Public Function Today() As Date:        Today = Date:                           End Function
Public Function Tomorrow() As Date:     Tomorrow = DateAdd("d", 1, Date):       End Function
Public Function Yesterday() As Date:    Yesterday = DateAdd("d", -1, Date):     End Function


Public Function FileCopy(SourceFile$, TargetFile$, Optional ErrMsg$ = "") As Boolean

    Dim FSO As Variant
    Dim Src As Variant
    Dim TRG As Variant

    On Error GoTo LocalError

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FileExists(SourceFile) Then
        If FileExists(TargetFile) Then
            Kill TargetFile
        End If
        Set Src = FSO.GetFile(SourceFile)
        Src.Copy TargetFile
        If FileExists(TargetFile) Then FileCopy = True
    End If
Exit Function

LocalError:
    ErrMsg = Err.Number & " - " & Err.Description
    FileCopy = False
End Function


Public Function isFileExist(Filename As String) As Boolean
  On Error GoTo FileDoesNotExist
  Call FileLen(Filename)
  FileExist = True
  Exit Function
FileDoesNotExist:
  FileExist = False
End Function


Public Function FileExists(ByVal Filename$) As Boolean

    Dim lngFileHandle As Long
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If ((Len(Filename) > 3) And (Right$(Filename, 1) = "\")) Then
        Filename = Left$(Filename, Len(Filename) - 1)
    End If
    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    FileExists = lngFileHandle <> INVALID_HANDLE_VALUE
    Call FindClose(lngFileHandle)

End Function

Public Function FileFind(RootPath$, Filename$) As String
    
    Dim lNullPos As Long
    Dim lResult As Long
    Dim sBuffer As String

    Const MAX_PATH = 260

    On Error GoTo LocalError

    'Allocate buffer
    sBuffer = Space(MAX_PATH * 2)
    'Find the file
    lResult = SearchTreeForFile(RootPath, Filename, sBuffer)

    'Trim null, if exists
    If lResult Then
        lNullPos = InStr(sBuffer, vbNullChar)
        If Not lNullPos Then
            sBuffer = Left$(sBuffer, lNullPos - 1)
        End If
        'Return filename
        FileFind = sBuffer
    Else
        'Nothing found
        FileFind = vbNullString
    End If
Exit Function

LocalError:
    FileFind = vbNullString
End Function

Public Function FilePart(FullPath As String, Optional WhichPart As FilePartTypes = FileNameOnly) As String

    If Len(FullPath) = 0 Then Exit Function

    Dim lArray As Variant
    Dim lSeperator As String

    lSeperator = "\"
    If InStr(FullPath, "/") > 0 Then lSeperator = "/"

    Select Case WhichPart
        Case FileExtOnly
            If InStr(FullPath, ".") Then
                lArray = Split(FullPath, ".")
                FilePart = lArray(UBound(lArray))
            End If
        Case FileNameOnly, FileNameAndExt
            lArray = Split(FullPath, lSeperator)
            FilePart = lArray(UBound(lArray))
            If WhichPart = FileNameOnly Then
                lArray = Split(FilePart, ".")
                FilePart = lArray(LBound(lArray))
            End If
        Case FilePathOnly
            Dim lFileName As String
            lFileName = FilePart(FullPath, FileNameAndExt)
            FilePart = Replace(FullPath, lFileName, "")
    End Select

End Function


Public Function FileKill(FileMask$, Optional OlderThan As Variant, Optional Prompt As Boolean) As Boolean

    On Error GoTo LocalError

    If Not IsMissing(OlderThan) Then
        If IsDate(OlderThan) Then
            Dim NextFile As String
            OlderThan = CDate(OlderThan)
            NextFile = Dir(FileMask)
            If Prompt And Len(NextFile) > 0 Then
                Dim lResponse As VbMsgBoxResult
                lResponse = MsgBox("Delete file(s) " & FilePart(FileMask, FileNameAndExt), vbYesNo + vbExclamation)
                If lResponse = vbNo Then Exit Function
            End If
            While Len(NextFile) > 0
                If FileLastAccessed(FileMask) < OlderThan Then
                    Kill NextFile   'Delete the file
                End If
                NextFile = Dir
            Wend
        End If
    Else    'Just do it
        Kill FileMask
    End If
    FileKill = True
Exit Function

LocalError:
    If Err.Number = 53 Then
        'file(s) was not found - continue
        FileKill = True
    Else
        FileKill = False
    End If
End Function

Public Function FileLastAccessed(ByVal Filename$) As Date

    Dim datFileCreationDate As Date
    Dim lngFileHandle As Long
    Dim udtSystemTime As SYSTEMTIME
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function

    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    Call FileTimeToSystemTime(udtWinFindData.ftLastAccessTime, udtSystemTime)
    datFileCreationDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour + FileAdjustTime, udtSystemTime.wMinute, udtSystemTime.wSecond)
    FileLastAccessed = datFileCreationDate
    Call FindClose(lngFileHandle)

End Function

Public Function FileLastModified(ByVal Filename$) As Date

    Dim datFileCreationDate As Date
    Dim lngFileHandle As Long
    Dim udtSystemTime As SYSTEMTIME
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function
    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    Call FileTimeToSystemTime(udtWinFindData.ftLastWriteTime, udtSystemTime)
    datFileCreationDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour + FileAdjustTime, udtSystemTime.wMinute, udtSystemTime.wSecond)
    FileLastModified = datFileCreationDate
    Call FindClose(lngFileHandle)

End Function

Public Function FileAttributes(ByVal Filename$) As String

    Dim lngFileAttributes As Long
    Dim strFileAttributeFlags As String

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function

    lngFileAttributes = GetFileAttributes(Filename)
    If lngFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then FileAttributes = FileAttributes + "D"
    If lngFileAttributes And FILE_ATTRIBUTE_ARCHIVE Then FileAttributes = FileAttributes + "A"
    If lngFileAttributes And FILE_ATTRIBUTE_SYSTEM Then FileAttributes = FileAttributes + "S"
    If lngFileAttributes And FILE_ATTRIBUTE_HIDDEN Then FileAttributes = FileAttributes + "H"
    If lngFileAttributes And FILE_ATTRIBUTE_READONLY Then FileAttributes = FileAttributes + "R"

End Function


Public Function FileRead(ByVal Filename$) As String

    Dim lngFileHandle As Long

    On Error Resume Next

    If FileExists(Filename) Then
        If Not InStr(FileAttributes(Filename), "D") Then
            lngFileHandle = FreeFile
            Open Filename For Binary As #lngFileHandle
            FileRead = Space(FileLen(Filename))
            Get #lngFileHandle, , FileRead
            Close #lngFileHandle
        End If
    End If

End Function

Public Function FileShortPath(ByVal Filename$) As String

    Dim strBuffer As String * 255
    Dim lngReturnCode As Long

    lngReturnCode = GetShortPathName(Filename, strBuffer, 255)
    FileShortPath = Left$(strBuffer, lngReturnCode)

End Function

Public Function FileSize(ByVal Filename As String) As Long

    'Get the file Size
    FileSize = FileLen(Filename) \ 1024

End Function

Public Function FileWrite(ByVal Filename$, ByVal FileContents$) As Boolean

    Dim lngFileHandle As Long

    On Error Resume Next

    If FileExists(Filename) Then
        If InStr(FileAttributes(Filename), "D") Then
            Exit Function
        Else
            Kill Filename
        End If
    End If

    lngFileHandle = FreeFile
    Open Filename For Binary As #lngFileHandle
    Put #lngFileHandle, , FileContents
    Close #lngFileHandle
    FileWrite = True

End Function

Private Function FileAdjustTime() As Long

    Dim datSystemDate As Date
    Dim udtSystemTime As SYSTEMTIME

    On Error Resume Next

    Call GetSystemTime(udtSystemTime)
    datSystemDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour, udtSystemTime.wMinute, udtSystemTime.wSecond)
    FileAdjustTime = DateDiff("h", datSystemDate, Now)

End Function

Public Function PathCheck(ByVal PathName$, Optional AltDelimiter$ = "") As String

    Dim Delimiter As String

    Delimiter = IIf(InStr(PathName, "/"), "/", "\")
    PathCheck = IIf(Right$(PathName, 1) = Delimiter, PathName, PathName & Delimiter)
    PathCheck = IIf(Len(AltDelimiter) = 0, PathCheck, Replace(PathCheck, Delimiter, AltDelimiter))

End Function

Public Function StripText(ByRef TextIN$, Optional Unwanted$)

    Dim currLoc As Integer
    Dim tmpChar As String

    If Len(Unwanted) = 0 Then Unwanted = "~`!@#$%^&*{}[]()_+-=|\?/.>,<" & Chr(34)

    For currLoc = 1 To Len(TextIN)
        tmpChar = Mid$(TextIN, currLoc, 1)
        If InStr(Unwanted, tmpChar) Then
            tmpChar = " " 'replace with a space
        End If
        StripText = StripText & tmpChar
    Next

    StripText = TrimALL(StripText)

End Function

Public Function TrimALL(ByVal TextIN As String) As String

    TrimALL = Trim(TextIN)

    While InStr(TrimALL, String(2, " ")) > 0
        TrimALL = Replace(TrimALL, String(2, " "), " ")
    Wend

End Function

Public Function TrimNull(InString As String) As String

    Dim Pos As Long

    Pos = InStr(InString, Chr$(0))
    TrimNull = IIf(Pos > 0, Left$(InString, Pos - 1), InString)
End Function

Public Sub UnloadForms(oForm As Form)

    Dim lForm As Form
    
    For Each lForm In Forms
        If lForm.Name <> oForm.Name Then
            Unload lForm
            Set lForm = Nothing
        End If
    Next lForm
    Set oForm = Nothing

End Sub

Function RetDate(Param) As String
    RetDate = Format(Param, "mm/dd/yyyy")
End Function

Function RetCurrency(Param) As String
    RetCurrency = Format(Param, "###,##0.00")
End Function


Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
' Example: Call FormOnTop(me.hWnd, True)
Dim wFlags, Placement
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub
