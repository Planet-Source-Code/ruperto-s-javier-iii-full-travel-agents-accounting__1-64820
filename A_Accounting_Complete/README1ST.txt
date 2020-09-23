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
'// 5.) Banking System (Bank Reconcilation) ( Credit / Debit )
'// 6.) Financial Reports
'// 7.) and so many others
'// 
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
'// * A/R Accounts Receivable with Twain Device Support
'//                                              - used by CyberLink Compu Sales
'// * Internet Cafe Time Management              - used by 10 internet cafe here in Iloilo Philippines
'// * Lending Services Software                  - used by RAK Lending Investor
'// * Savings / Time Deposit                     - used by the Bank were i was employed now
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

username = img <- has system wide privileges
password = x


logon to http://www.imaginasoft.co.nr for OCX download