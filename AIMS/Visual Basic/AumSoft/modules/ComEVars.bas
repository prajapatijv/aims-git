Attribute VB_Name = "ComEVars"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B83007701F2"
Option Explicit

Declare Function CreateFieldDefFile Lib "p2smon.dll" (X As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer
Declare Function CreateReportOnRuntimeDS Lib "p2smon.dll" (X As Object, ByVal reportPath$, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%, ByVal bLaunchDesigner%) As Integer

Public Declare Function PS Lib "winmm.dll" Alias "PlaySound" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'Used to Add Fonts
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global gBackEnd As BackEnd_Type
Global gPathReport As String
Global gPathInv As String
Global gPathResources As String
Global gBackUpPath As String
Global gBgImagePath As String

Global gMdbMst As String
Global gDsn_Name As String
Global gPrintEnable As Boolean

Global gDefaComp As Integer

Global gUser As String
Global gUserName As String
Global gPasswd As String
Global gOwner As String

Global gConnectStr As String
Global gConnectStrDbf As String
Global gDefaLockEdit As LockTypeEnum

Global gDateFormat 'As String = "dd/mm/yyyy"
Global gMdbPwd As String

Global gCnnMst As ADODB.Connection

Global gSrvName As String
Global gSrvUID As String
Global gSrvPwd As String
Global gTerminalId As Integer
Global gAIMS_SERVER_TYPE As String

Global gAdd As Boolean
Global gEdit As Boolean
Global gDel As Boolean
Global gPrint As Boolean
Global gNevigate As Boolean
Global gTktFmt As Integer

Global gGujaratiFontName As String
Global SQL As String
Global gDenyZeroPriceMaterialInwardOutwardTypes As String
Global gShowTerminalCodeForInwardOutwardTypes As String

'Global FromWhichForm As Form
'Global gReportTag As String

Public Enum BackEnd_Type
    BE_Access = 0
    BE_Oracle = 1
    BE_SQLSrv = 2
    BE_CryRep = 3
End Enum

Public Enum Lan_LockTypes
    ReadRecLock = 3624 'Couldn't read the record currently locked by another user.
    UpdLock = 3260 'Couldn't update currently lock by user .... on machine ....
    ReadLock = 3187 'Couldn't read currently lock by user .... on machine ....
    FileLock = 3050 'Couldn't Lock File
    SameData = 3197 'you & another user are attempting to change the same data at same time
    FileExclOpen = 3051 'file exclusive open
    TabExlcLock = 3261 'Table ______ is exclusive locked by user _______  on machine ____ TooManyUsers = 3239 'Too many active users.
    ExclDBOpen = 3356 'You try to open the database that is already opened exlcusive by user .......... on machine .......... try again whene the database is avalabile.
    TooManyUsers = 3239 'Too Many Active Users.
'Duplicate Values
    DuplValues = 3022 'The changes you requested to the table would create duplicate values.
End Enum

Public Enum ButtonType
    btnDash1 = 1
    btnadd = 2
    btnedit = 3
    btndel = 4
    btnDash2 = 5
    btnsave = 6
    btnSaveNAdd = 7
    btnCancel = 8
    btnDash3 = 9
    btnprint = 10
    btnview = 11
    btnDash4 = 12
    btnfirst = 13
    btnnext = 14
    btnprev = 15
    btnlast = 16
    btnDash7 = 17
    btnFind = 18
    btnDash5 = 19
    btnExit = 20
    btnDash6 = 21
End Enum

Public Enum Nevigate
    MoveFirst = 1
    MoveNext = 2
    MovePrev = 3
    MoveLast = 4
    MoveTo = 5
End Enum

Public Enum ViewMode
    EntryReadWrite = 1
    EntryReadOnly = 2
End Enum

Public Enum enOperationMode
    enServer = 1
    enTerminal = 2
End Enum


