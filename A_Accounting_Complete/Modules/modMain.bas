Attribute VB_Name = "modMain"
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//ELS Travel and Tours ( Multi-User )
'//
'//MODULES INCLUDED
'//
'// 1.) Accounting
'// 2.) Payments ( Multiple )
'//         - Cash,Check,Credit Cards, Others
'//         - Can Handle Dollar,Peso (Multi-Currency)
'// 3.) Purchase Order
'// 4.) Customer Ledger
'// 5.) Banking System (Bank Reconcilation)
'// 6.) Financial Reports
'// 7.) and so many others
'//
'//
'//
'// PROGRAMMER : RUPERTO S. JAVIER III (boykulot)
'// ADDRESS    : #23 VALERIA ST., ILOILO CITY
'// EMAIL      : imagina_boy@eudoramail.com
'//              imagina_boy@linuxmail.org
'// URL        : http://www.imaginasoft.co.nr ( under development )
'// CEL #      : +639213378103
'//
'// SOME OTHER PROGRAMS FINISHED FULLY FUNCTIONAL AND WERE USED BY THE COMPANY
'// I CONTRACTED :
'//
'// * POS (Point of Sale) for Pharmacy & Grocery - currently being used in Pharmacy here in Iloilo Philippines
'// * Inventory for Domescon Motors              - used by DOMESCON MOTORS
'// * Inventory for Cyberlink Compu Sales        - used by CyberLink Compu Sales
'//   Barcode fully supported
'// * A/R Accounts Receivable with Twain Device Support
'//                                              - used by CyberLink Compu Sales
'// * Internet Cafe Time Management              - used by 10 internet cafe here in Iloilo Philippines
'// * Lending Services Software                  - used by RAK Lending Investor
'// * Savings / Time Deposit w/ Passbook Printer - used by the Bank were i was employed now
'// * Interfacing VB
'//
'//
'//
'// Tools Used in my project
'//     1.) Visual Basic 6.0
'//     2.) Visual Foxpro
'//     3.) MS Access for small client
'//     4.) mySQL for multi-branch
'//     5.) For Web I used ASP and PHP mix of both
'//     6.) Sometime used flash for nice GUI
'//     7.) My Development OS was winXP pro
'//     8.) Used VMWare Workstation for my Linux Development
'//     9.) My FileServer for LAN was linux using SAMBA i used TRUSTIX Distro very nice
'//     10.) For Reports i heavily used ACTIVE REPORTS very handy
'//
'//
'// CREDITS :
'//             THE AUTHOR OF LAVOLPE BUTTON    - his button was heavily used on this project
'//             THE AUTHOR OF CLSHUFFMAN        - just for compression very nice
'//             THE AUTHOR OF UCGRADCONTAINER   - some forms were skinned with this control
'//             THE AUTHOR OF CLSBUSY           - im too lazy to code thnx for ur very simple class
'//             THE AUTHOR OF HOOKMENU          - VERY GOOD
'//             AND SO MANY OTHERS              - some code were borrowed from
'//                                               authors which i forgot
'//
'//
'// A NOTE OF WARNING!!!!
'//         SOME CODES/VALUES WERE HARD-CODED
'//         SOME COMPUTATION WERE DONE IN QUERY
'//         HEAVILY USED RELATIONSHIP IN MSACCESS - try to view the relationship
'//         RELY ON EXTERNAL FILE FOR GENERATION OF AUTONUMBER
'//         SO MANY FORMS very Difficult to debug ( CAN BE OPTIMIZED IF I HAVE TIME )
'//         SOME CODES CAN BE FOUND IN OTHER MODULES (TOO LAZY TO OPTIMIZED)
'//         THIS PROJECT IS QUIET HUGE!!!! TRY TO UNDERSTAND FIRST BEFORE EDITING
'//
'// FINAL WORDS :
'//         USED THIS PROJECT AS YOUR BASIS IN BUILDING PROJECT
'//         IF U PLAN TO USE THIS PROJECT INFORM ME ( I CAN HELP )
'//         I STRONGLY BELIEVE IN OPEN SOURCE SOFTWARE ( THATS WHY IM GIVING YOU MY CODE )
'//         IF YOU HAVE ANY PROJECTS AND NEEDS HELP PLS DONT HESITATE TO GIVE A NOTICE I MAY HELP
'//         THIS PROJECT IS 98% WORKING!!!!
'//
'// I AM LOOKING FOR PROGRAMMING JOB FULL-TIME EMAIL ME AT THE ADDRESS ABOVE
'//
'// THIS IS MY FIRST SUBMISSION PLS.. VOTE
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'*************************************************************************
'* You are welcome to use this in your projects, as long as              *
'* comments containing all people named in the credits remain intact.    *
'* Only LAME codeRZ thieves download code, remove                        *
'* the comments, and claim they wrote it.  If you are distributing this  *
'* as part of a compiled application and the application has a 'Credits' *
'* section, named credit is appreciated, but not required.               *
'*************************************************************************


Option Explicit
Public cn                                   As ADODB.Connection
Public cnLocal                              As ADODB.Connection
Global Const CryptPass = "272:8816"
Public MSDatabase                           As String
Public UserDatabase                         As String


'change the Following for your own path to DB

'Global Const NetWorkPath = "\\dbserver\DATA"
Global Const NetWorkPath = "G:\ELScode\els\Data"
'Global Const NetWorkPath = "d:\els\Data"
Global Const AppCaption = "ELS TRAVEL and TOURS (#113 Ledesma St., Iloilo City Philippines"

Public dB                                   As ADODB.Connection
Public DBConnect                            As Boolean
Public UserConnected                        As Boolean
Public WhichBranch                          As ADODB.Recordset
Public RsBackUpSettings                     As ADODB.Recordset

Public oRegistry                            As New Registry                 ' Registry Class
Public oLoader                              As New clsBusy


'Global variables
Public myPCnum                              As Variant
Public myGlobal_LilyHo_AccNo                As String
Public myGlobal_PAL_MBTC                    As String
Public myGlobal_CP_MBTC                     As String
Public myGlobal_AP_AS_EPCI                  As String
Public myGlobal_CK_NN_WGA_TA_EPCI           As String
Public myGlobal_PAL_HSBC                    As String
Public myGlobal_DOLLAR_ACC                  As String

'myGlobal Airlines
Public myGlobal_PAL                         As String
Public myGlobal_CP                          As String
Public myGlobal_CK                          As String
Public myGlobal_NN                          As String
Public myGlobal_AP                          As String
Public myGlobal_AS                          As String
Public myGlobal_WGA                         As String
Public myGlobal_TA                          As String

'For Report Company

Public myGlobal_CashDue                     As Double
Public myGlobal_RefundAmt                   As Double
Public myGlobal_SalesAmt                    As Double
Public myglobal_NetSales                    As Double



Dim STRSQL As String

Public Function DataConnect() As Boolean
On Error GoTo OpenErr

oLoader.BusyStatus 0, "Connecting please wait...."
    DoEvents 'lets you execute methods
    Set cn = New ADODB.Connection
        cn.CursorLocation = adUseClient
    Set cnLocal = New ADODB.Connection
        cnLocal.CursorLocation = adUseClient
        MSDatabase = NetWorkPath & ("\dbMaster.MDB")
        
        cn.CursorLocation = adUseClient
        cn.Provider = "Microsoft.Jet.OLEDB.4.0; Jet OLEDB:Database Password=" & Decrypt(CryptPass)
        cn.Open MSDatabase ', Admin
        Display_Other_Task
frmLoading.Show 1

DataConnect = True
Exit Function
OpenErr:
    Select Case Err.Number
    Case -2147467259
            DataConnect = False
            oLoader.BusyExit
    Case Else
    End Select
    
            DBConnect = False

End Function

Sub Main()
Dim ask As Integer

  DBConnect = True
        Call DataConnect
    If DBConnect = True Then
            Set RsBackUpSettings = New ADODB.Recordset
            RsBackUpSettings.Open "SELECT * FROM tbl_BackUpSettings", cn, adOpenKeyset, adLockOptimistic
            Set WhichBranch = New ADODB.Recordset
            With WhichBranch
                  .Open "SELECT * FROM tbl_SetBranch", cn, adOpenKeyset, adLockOptimistic
            End With
            
'//Initialize Values
'           /for Bank Account Grouping
            myGlobal_PAL_MBTC = "PAL-MBTC"
            myGlobal_CP_MBTC = "CP-MBTC"
            myGlobal_AP_AS_EPCI = "AP/AS-EPCI"
            myGlobal_CK_NN_WGA_TA_EPCI = "CK/NN/WGA/TA-EPCI"
            myGlobal_PAL_HSBC = "PAL-HSBC"
            myGlobal_DOLLAR_ACC = "8888888888888"
            
'           /For airlines
            myGlobal_PAL = "PHIL-AIRLINE"
            myGlobal_CP = "CEBU-PAC"
            myGlobal_CK = "COKALIONG"
            myGlobal_NN = "NN"
            myGlobal_AP = "AP"
            myGlobal_AS = "ASIAN SPIRIT"
            myGlobal_WGA = "WGA"
            myGlobal_TA = "TRANS-ASIA"
            
            MDImain.Show
            frmDateChecker.Show 1
    Else
        MsgBox "Network connection error! Please verify that the connection is good!" & _
               Chr(13) & Chr(13) & "Possible cause :" & _
               Chr(13) & "Network down, corrupt database, cable connection error", vbCritical, "imgSoft"
        ask = MsgBox("Reconnect to server?", vbYesNo + vbInformation)
        If ask = vbYes Then
            Call Main
        End If
    End If

End Sub

Sub Display_Other_Task()
On Error GoTo ErrExit
oLoader.BusyStatus 5, "Initializing..."
Sleep 250

'+++++++++++++++++++++++++++++++++++++++++++++++++++++
'Now Load external files for processing
'   1.) Settings.txt was used for statement number
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
oLoader.BusyStatus 10, "Loading external files..."

    If isFileExist(App.Path & "\Settings.txt") Then
        myPCnum = kulotRead(App.Path & "\Settings.txt")
        oLoader.BusyStatus 15, "Settings successfully loaded"
    Else
        oLoader.BusyExit
        'GoTo ErrExit
    End If
    
Sleep 250       '   just delay 250 milliseconds about 1/4 a sec

'+++++++++++++++++++++++++++++++++++++++++++++++++++++
'Now Load plug-in files if there are any
'   1.)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
oLoader.BusyStatus 60, "Loading Plug-in..."
Sleep 250


oLoader.BusyStatus 100, "Completing Startup..."
Sleep 250

oLoader.BusyExit

Exit Sub

ErrExit:
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
'if errors exist never allow the app to continue...
'just quit and restart
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
MsgBox "There was an error while starting up the application quiting...", vbInformation
End
End Sub
