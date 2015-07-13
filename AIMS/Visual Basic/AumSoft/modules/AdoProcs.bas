Attribute VB_Name = "AdoProcs"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B8300CF01E4"
Option Explicit

'Public Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "User32" () As Long
Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long

Private DBname As String
Private mMdbpath As String
    
Public Sub OpenAdoRst(s_AdoRst As ADODB.Recordset, s_Source, Optional s_CursorType As ADODB.CursorTypeEnum = adOpenDynamic, Optional s_CmdType As CommandTypeEnum = adCmdText, Optional s_LckType As LockTypeEnum = adLockOptimistic, Optional s_Cnn As ADODB.Connection)

    If s_Cnn Is Nothing Then
        Set s_Cnn = gCnnMst
    End If
    
    Set s_AdoRst = New ADODB.Recordset
    s_AdoRst.Open s_Source, s_Cnn, s_CursorType, s_LckType
    
End Sub

Public Sub CloseAdoRst(s_AdoRst As ADODB.Recordset, Optional s_Nothing As Boolean = True)
    If s_AdoRst Is Nothing Then Exit Sub
    If s_AdoRst.State = adStateOpen Then s_AdoRst.Close
    If s_Nothing Then
        Set s_AdoRst = Nothing
    End If
End Sub


Public Sub SetFormByRes(s_Frm As Form)
Dim ctlTemp As Control

Dim mMultifact As Double
Dim mFntSize As Integer
Dim mDontSet As String

If GetScrRes = "800 X 600" Then
    Exit Sub
ElseIf GetScrRes = "1024 X 768" Then
    mMultifact = 1.28
    mFntSize = 12
End If

'Exit Sub

mDontSet = "timer~statusbar~menu~line~"

For Each ctlTemp In s_Frm.Controls
    
    With ctlTemp
        'Debug.Print LCase(TypeName(ctlTemp))
    Select Case LCase(TypeName(ctlTemp))
        Case "frame"
            .Height = .Height * mMultifact
            .Width = .Width * mMultifact '+ 25
            .Left = (.Left * mMultifact) '+ 30
            .Top = (.Top * mMultifact)
        
        Case "combobox", "textbox", "label", "dtpicker", "commandbutton", "msflexgrid", "maskedbox", "datacombo", "listview", "sstab"
            If LCase(TypeName(ctlTemp)) <> "combobox" Then .Height = (.Height * mMultifact) '- 25
            .Width = (.Width * mMultifact) ' - 25
            .Left = (.Left * mMultifact)
            .Top = (.Top * mMultifact)
            'If LCase(TypeName(ctlTemp)) <> "dtpicker" Then .FontSize = mFntSize
            
        Case "hnsctxt", "hnsitxt", "hnsntxt", "hnstxtbox", "hnsntxtbox"
            If LCase(TypeName(ctlTemp)) <> "combobox" Then .Height = (.Height * mMultifact) '- 25
            .Width = (.Width * mMultifact) ' - 25
            .Left = (.Left * mMultifact)
            .Top = (.Top * mMultifact)
            'If LCase(TypeName(ctlTemp)) <> "dtpicker" Then .FontSize = mFntSize
        Case "line"
            .X2 = .X2 * mMultifact
            .Y1 = (.Y1 * mMultifact)
            .Y2 = (.Y2 * mMultifact)

        
        Case "timer", "statusbar", "menu", "adodc"
        
    End Select
    Select Case LCase(TypeName(ctlTemp))
        Case "timer", "statusbar", "menu", "adodc", "sstab"
        Case "hnsctxt", "hnsitxt", "hnsntxt", "hnstxtbox", "hnsntxtbox", "hnscombo"
        Case Else
            '.Refresh
    End Select
    End With
    
Next

End Sub

Public Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   End If
   GetActiveWindow = i
End Function

Public Sub ExpSql2Mdb(s_MdbTable As String, s_SqlTable As String, Optional s_fldMach As Boolean = False, Optional s_Where As String = "")

Dim rst As ADODB.Recordset
Dim Db As DAO.Database
Dim tdf As DAO.TableDef
Dim i As Integer

Dim mStr As String
Dim mMdbFlds As String, mSqlFlds As String

Dim gConnectStr As String
Dim WsMdb As DAO.Workspace
Dim DbMdb As DAO.Database
Dim rstMdb As DAO.Recordset
    
    
    gConnectStr = ";PWD="
    
    mMdbpath = Readfromini("ExportMdbPath", App.Path + "\" + SetPkgName)
    If OperaionMode = enServer Then
        DBname = mMdbpath & "ServerImpex.mdb"
    Else
        DBname = mMdbpath & "TerminalImpex" & Trim$(gTerminalId) & ".mdb"
    End If
    
    Set WsMdb = DBEngine.Workspaces(0)
    
    If Trim$(Dir(DBname)) = "" Then
        WsMdb.CreateDatabase DBname, dbLangGeneral, dbEncrypt
    End If
    
    Set DbMdb = WsMdb.OpenDatabase(DBname, True, False, gConnectStr)
    
    If s_fldMach Then
        Set rstMdb = DbMdb.OpenRecordset(s_MdbTable)
        mMdbFlds = "/"
        For i = 0 To rstMdb.Fields.Count - 1
            mMdbFlds = mMdbFlds & rstMdb.Fields(i).Name & "/"
        Next
        
        OpenAdoRst rst, s_SqlTable, adOpenDynamic, adCmdTable, dbPessimistic
        mSqlFlds = ""
        For i = 0 To rst.Fields.Count - 1
            mStr = "/" & rst.Fields(i).Name & "/"
            If InStr(1, mMdbFlds, mStr) > 0 Then
                mSqlFlds = mSqlFlds & rst.Fields(i).Name & ","
            End If
        Next
        mSqlFlds = Mid(mSqlFlds, 1, Len(mSqlFlds) - 1)
        SQL = "insert into " & s_MdbTable & " IN '" & DBname & "'" & " select " & mSqlFlds & " from " & s_SqlTable & s_Where
    Else
        If TableExists(s_MdbTable, DBname, gConnectStr) Then
            DbMdb.Execute "drop table " & s_MdbTable
        End If
        SQL = "select * into [;database=" & DBname & gConnectStr & "]." & s_MdbTable & " from " & s_SqlTable
        SQL = SQL & s_Where
    End If
    
    Set Db = OpenDatabase("", False, False, "ODBC;DSN=" & gDsn_Name & ";Uid=" & gSrvUID & ";Pwd=" & gSrvPwd & ";")
    Db.Execute SQL
    
    On Error GoTo errorhandler
    
exit_ImpData:
    Exit Sub
    
errorhandler:
    MsgBox Err.Number & vbCrLf & Err.Description
    GoTo exit_ImpData
    
End Sub


Private Function ConvDatForAccess(s_text As String) As String
    ConvDatForAccess = "#" & Format(s_text, "mm/dd/yyyy") & "#"
End Function

Public Sub DeleDataFromMdb(s_Table As String, s_Cond As Variant)
    Dim Ws As Workspace
    Dim Db As Database
    Dim rst As Recordset
    Dim delSQL As String
    Dim gConnectStr As String
    Dim mMdbpath As String
    
    gConnectStr = ";PWD="
    
    mMdbpath = Readfromini("ExportMdbPath", App.Path + "\" + SetPkgName)

    DBname = mMdbpath & "ServerImpex.mdb"
    
    Set Ws = DBEngine.Workspaces(0)
    Set Db = Ws.OpenDatabase(DBname, True, False, gConnectStr)
    Set rst = Db.OpenRecordset(s_Table)
    
    Ws.BeginTrans
        delSQL = "DELETE * FROM " & s_Table & _
               " WHERE " & s_Cond
        
        Db.Execute delSQL
    Ws.CommitTrans
    
    rst.Close
    Db.Close
    Ws.Close
    Set rst = Nothing
    Set Db = Nothing
    Set Ws = Nothing
End Sub


Private Function TableExists(TableName As String, DBname As String, Optional f_ConStr As String) As Boolean
Dim mConStr As String
Dim Db As Database
Dim i As Integer
Dim mRetVal As Boolean

If TLen(f_ConStr) = 0 Then
    mConStr = gConnectStr
Else
    mConStr = f_ConStr
End If

Set Db = DBEngine.Workspaces(0).OpenDatabase(DBname, False, False, mConStr)

For i = 0 To Db.TableDefs.Count - 1
    If LCase(Db.TableDefs(i).Name) = LCase(Trim(TableName)) Then
        mRetVal = True
        Db.Close
        Set Db = Nothing
        TableExists = mRetVal
        Exit Function
    Else
        mRetVal = False
    End If
Next

TableExists = mRetVal

Db.Close
Set Db = Nothing

Exit Function
Lc:
    MsgBox " TableExists "
    Resume Next
End Function

Public Sub ImportFromMdb2Sql(s_TableName As String, s_MdbName As String)

    mMdbpath = Readfromini("ImportMdbPath", App.Path + "\" + SetPkgName)
    
    If OperaionMode = enServer Then
        DBname = mMdbpath & s_MdbName
        
        Select Case LCase(s_TableName)
            Case LCase("TerSaltrn")
                SQL = " Exec stpImportFromMdb2Sql_Saltrn " & AQ(DBname)
                
            Case Else
                'Nothing
        End Select
        
        gCnnMst.Execute SQL
    Else
    
        DBname = mMdbpath & s_MdbName
        
        SQL = " Exec stpImportFromMdb2Sql " & AQ(DBname) & "," & _
                                                s_TableName & "," & _
                                                1                       '1 - Truncate table b4 Importing
        gCnnMst.Execute SQL
    End If

Exit Sub
Lc:
    MsgBox Err.Number & vbCrLf & Err.Description
    Resume Next
End Sub
