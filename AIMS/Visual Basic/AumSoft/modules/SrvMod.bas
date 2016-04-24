Attribute VB_Name = "SrvMod"
Option Explicit

Public Sub Main()
    SetMainVars
End Sub

Public Sub SetMainVars()
If App.PrevInstance = True Then End
Dim mIniFile As String
Dim IsIniFound As Boolean
    
    gOwner = ".dbo."
    
    mIniFile = App.Path + "\" + SetPkgName

    mIniFile = Trim(App.Path) & "\" & SetPkgName + ".ini"
    
    IsIniFound = True
    If Not IsIniFound Then
        MsgBox "Ini Settings File not Found", vbExclamation
        End
    End If
    
    CreateIniFile

    If Not CheckDateFormat Then End

    
    gTerminalId = Readfromini("TerminalId", mIniFile)
    gGujaratiFontName = "Kanaiya-Normal" '"GujaFont"
    gBackUpPath = Readfromini("BackUpPath", mIniFile)

    gSrvName = Readfromini("servername", mIniFile)
    gMdbMst = Readfromini("mstdb", mIniFile)
    gDsn_Name = Readfromini("Dsn_Name", mIniFile)
    gPathReport = Readfromini("ReportPath", mIniFile)
    gBgImagePath = Readfromini("BgImagePath", mIniFile)
    gPrintEnable = Readfromini("PrintEnable", mIniFile)
    gTktFmt = Readfromini("TktFmt", mIniFile)
    gAIMS_SERVER_TYPE = Readfromini("AIMS_SERVER_TYPE", mIniFile)
    gDenyZeroPriceMaterialInwardOutwardTypes = Readfromini("DenyZeroPriceMaterialInwardOutwardTypes", mIniFile)
    gShowTerminalCodeForInwardOutwardTypes = Readfromini("ShowTerminalCodeForInwardOutwardTypes", mIniFile)

    gDocumentPath = Readfromini("DocumentPath", mIniFile)
    gDocumentTypesFilter = Readfromini("DocumentTypesFilter", mIniFile)
    
    gSrvUID = "sa"
    gSrvPwd = ""
    gMdbPwd = ""
    gConnectStr = ";PWD=" & gMdbPwd
    
    
    gBackEnd = BE_SQLSrv
    
    If gBackEnd = BE_SQLSrv Then gDateFormat = "mm/dd/yyyy"

    Set gCnnMst = New ADODB.Connection
    
    gCnnMst.ConnectionTimeout = 5
    gCnnMst.CommandTimeout = 0
    gCnnMst.CursorLocation = adUseClient
    gCnnMst.Open GetCnnStrMst(gDsn_Name)
              
    CreatePkgMstTables
    AddPkgConstraintsMst
    
    frmSplash.Show
    Exit Sub

End Sub

Public Function GetCnnStrMst(Optional s_DsnName As String = "dsn_mst") As String
    GetCnnStrMst = s_DsnName
    'GetCnnStrMst = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=" & gSrvUID & ";password=" & gSrvPwd & ";Initial Catalog=" & gMdbMst & ";Data Source=" & gSrvName
End Function

Public Function GetCnnStrSql(f_SrvUID As String, f_SrvPwd As String, f_Database As String, f_SrvName As String) As String
    If gBackEnd = BE_SQLSrv Then GetCnnStrSql = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & f_SrvUID & ";password=" & f_SrvPwd & ";Initial Catalog=" & f_Database & ";Data Source=" & f_SrvName
End Function


