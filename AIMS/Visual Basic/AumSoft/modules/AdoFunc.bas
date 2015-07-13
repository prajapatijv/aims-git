Attribute VB_Name = "AdoFunc"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B8300CB03AB"
Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400

Public Function AdoIsTable(s_TableName As String, Optional s_Cnn As ADODB.Connection) As Boolean
Dim rst As ADODB.Recordset
Dim mRtnVal As Boolean

    SQL = "select name from sysobjects where name='" & s_TableName & "'"
    
    Set rst = New ADODB.Recordset
    
    OpenAdoRst rst, SQL, , , , s_Cnn
    
    If rst.RecordCount > 0 Then
        mRtnVal = True
    Else
        mRtnVal = False
    End If
    
    rst.Close
    Set rst = Nothing
    AdoIsTable = mRtnVal

End Function

Public Function AdoIsDatabase(s_Db As String) As Boolean
Dim rst As ADODB.Recordset
Dim mRtnVal As Boolean

    SQL = "select name from master.dbo.sysDatabases where name='" & s_Db & "'"
    
    Set rst = New ADODB.Recordset
    
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    If rst.RecordCount > 0 Then
        mRtnVal = True
    Else
        mRtnVal = False
    End If
    
    rst.Close
    Set rst = Nothing
    AdoIsDatabase = mRtnVal

End Function

Public Function AdoIsIndex(f_Indexname As String, Optional f_Cnn As ADODB.Connection, Optional f_reindex As Boolean = False) As Boolean
Dim rst As ADODB.Recordset
Dim Rtnvalue As Boolean
Dim mTblName As String

    If f_Cnn Is Nothing Then Set f_Cnn = gCnnMst

    If gBackEnd = BE_SQLSrv Then
        SQL = "select id,name,status "
        SQL = SQL & " from sysindexes "
        SQL = SQL & " Where name=" & AQ(f_Indexname)
        SQL = SQL & " and (status<>0 or indid <> 0)"
    End If
    
    Set rst = New ADODB.Recordset
    
    OpenAdoRst rst, SQL
    If rst.RecordCount > 0 Then
        
        If gBackEnd = BE_SQLSrv Then
            If f_reindex = True Then
                SQL = " select name from sysobjects "
                SQL = SQL & " Where id = " & rst.Fields("id").Value

                If f_Cnn.Execute(SQL).RecordCount > 0 Then
                    mTblName = Trim(f_Cnn.Execute(SQL).Fields("name").Value)
                    f_Cnn.Execute ("Drop index " & mTblName & "." & f_Indexname)
                    AdoIsIndex = False
                    rst.Close
                    Set rst = Nothing
                    Exit Function
                End If
            End If
        End If
                
        Rtnvalue = True
    
    Else
        Rtnvalue = False
    End If
    
    rst.Close
    Set rst = Nothing
    
    AdoIsIndex = Rtnvalue

End Function

Public Function AdoIsSP(s_SPName As String) As Boolean
Dim rst As ADODB.Recordset
Dim Rtnvalue As Boolean

    If gBackEnd = BE_Oracle Then
        SQL = ""
    ElseIf gBackEnd = BE_SQLSrv Then
        If LCase(Left(s_SPName, 3)) = "sp_" Then SQL = "SELECT * FROM SYSOBJECTS WHERE TYPE='P' AND NAME=" & AQ(Mid(s_SPName, 4, 100))
        If LCase(Left(s_SPName, 3)) = "tr_" Then SQL = "SELECT * FROM SYSOBJECTS WHERE TYPE='TR' AND NAME=" & AQ(Mid(s_SPName, 4, 100))
        If LCase(Left(s_SPName, 3)) = "fn_" Then SQL = "SELECT * FROM SYSOBJECTS WHERE TYPE='FN' AND NAME=" & AQ(Mid(s_SPName, 4, 100))
    End If
    
    Set rst = New ADODB.Recordset
    
    OpenAdoRst rst, SQL
    
    If rst.RecordCount > 0 Then
        Rtnvalue = True
    Else
        Rtnvalue = False
    End If
    
    rst.Close
    Set rst = Nothing
    AdoIsSP = Rtnvalue

End Function
Public Function GetScrRes() As String
    GetScrRes = Trim(str(Screen.Width / Screen.TwipsPerPixelX)) & " X " & Trim(str(Screen.Height / Screen.TwipsPerPixelY))
End Function

Public Function AdoIsConstraint(f_Cnst As String, Optional f_Cnn As ADODB.Connection) As Boolean
Dim RstIsCnst As ADODB.Recordset
Dim mRetVal As Boolean
Dim i As Integer

    SQL = "select constraint_name from "
    SQL = SQL & " INFORMATION_schema.TABLE_CONSTRAINTS "
    SQL = SQL & " where upper(constraint_name)=" & AQ(UCase(f_Cnst))
    
    OpenAdoRst RstIsCnst, SQL, , , , f_Cnn
    If RstIsCnst.RecordCount = 0 Then
        AdoIsConstraint = False
    Else
        AdoIsConstraint = True
    End If
End Function

Public Sub Shell32BitOld(ByVal JobToDo As String)

         Dim hProcess As Long
         Dim RetVal As Long
         'The next line launches JobToDo as icon,

         'captures process ID
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, 1))

         Do
             'Get the status of the process
             GetExitCodeProcess hProcess, RetVal

             'Sleep command recommended as well as DoEvents
             DoEvents: Sleep 100

         'Loop while the process is active
         Loop While RetVal = STILL_ACTIVE


End Sub

Public Function GetDbTable(s_TblName As String, Optional s_DbName As String) As String
Dim mRetVal As String
    If Len(s_DbName) <> 0 Then
        mRetVal = " " & s_DbName & gOwner & s_TblName & " "
    Else
        mRetVal = " " & gMdbMst & gOwner & s_TblName & " "
    End If
    
    GetDbTable = mRetVal
End Function

Public Function AdoIsHasTable(s_TblName As String, Optional s_Cnn As String = "sale") As Boolean
On Error GoTo errhndl

    Dim rst As ADODB.Recordset
    Dim SQL As String
    
    SQL = "select name from tempdb.dbo.sysobjects where id = object_id('tempdb.dbo." & Trim(s_TblName) & "') "
    OpenAdoRst rst, SQL
    
    If Not rst.EOF Then
        AdoIsHasTable = True
    Else
        AdoIsHasTable = False
    End If
    
Exit Function
errhndl:
    'Errmsg
    Resume Next
End Function

Public Function MonthDays(s_month As Integer, s_year As Integer) As Integer

    Select Case s_month
        Case 1, 3, 5, 7, 8, 10, 12
            MonthDays = 31
        Case 4, 6, 9, 11
            MonthDays = 30
        Case 2
            If s_year Mod 4 = 0 Then
                MonthDays = 29
            Else
                MonthDays = 28
            End If
    End Select
    
End Function

Public Function GetMonthInitialDat(s_Dat As String)
    
    GetMonthInitialDat = "01/" & Format(Month(s_Dat), "00") & "/" & Year(s_Dat)
    
End Function

Public Function GetFieldWidth(f_tblname As String, f_FldName As String) As Integer
Dim mLen As Integer
    
    mLen = 0
    
    SQL = " Select length from syscolumns "
    SQL = SQL & " Where id IN"
    SQL = SQL & " (select id from sysobjects "
    SQL = SQL & " Where type='u' "
    SQL = SQL & " and name=" & AQ(f_tblname) & ")"
    SQL = SQL & " and name=" & AQ(f_FldName)
    mLen = gCnnMst.Execute(SQL).Fields("length").Value
    
    GetFieldWidth = mLen
    
End Function

Public Sub Shell32Bit(ByVal JobToDo As String, Optional s_WinFocus As Integer = 1)

         Dim hProcess As Long
         Dim RetVal As Long
         'The next line launches JobToDo as icon,

         'captures process ID
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, s_WinFocus))

         Do
             'Get the status of the process
             GetExitCodeProcess hProcess, RetVal

             'Sleep command recommended as well as DoEvents
             DoEvents: Sleep 100

         'Loop while the process is active
         Loop While RetVal = STILL_ACTIVE

End Sub

Public Function GetMaxCode(f_TableName As String, Optional f_Short As Boolean = False, Optional f_CodeField As String = "Code", Optional s_Cnn As ADODB.Connection, Optional f_AddCond As String = "") As Long

    If s_Cnn Is Nothing Then s_Cnn = gCnnMst
    
    SQL = " Select isnull(Max(" & f_CodeField & ")," & IIf(f_Short, 1000, 100000) & ") as MaxCode "
    SQL = SQL & " from " & f_TableName
    SQL = SQL & " Where 1=1 "
    SQL = SQL & " " & f_AddCond
        
    GetMaxCode = s_Cnn.Execute(SQL).Fields("MaxCode").Value + 1
    
End Function

Public Function GetInsertFieldStr(f_TableName As String) As String

    SQL = " Select isnull([name],'')  as FldName from "
    SQL = SQL & " SysColumns "
    SQL = SQL & " Where id = ( Select id from SysObjects "
    SQL = SQL & " Where [name] = " & AQ(f_TableName) & ")"

    GetInsertFieldStr = gCnnMst.Execute(SQL).GetString
    
End Function

Public Function GetMaxVno(f_TableName As String, Optional s_Where As String = "") As Long
    SQL = " Select isnull(Max(Vno),0) as MaxVno "
    SQL = SQL & " From " & f_TableName
    SQL = SQL & s_Where
    
    GetMaxVno = gCnnMst.Execute(SQL).Fields("MaxVno").Value + 1
End Function

Public Sub AdoRsRead(rs As ADODB.Recordset)
On Error GoTo errhndl
Dim i As Integer
Dim ctlTemp As Control
Dim strFldname As String
Dim mFldList As String  'Var defined to store field names
Dim mValidObjs As String

    If rs.EOF Or rs.BOF Then Exit Sub
    
    'Store FieldNames to Var (mFldlist)
    mFldList = ""
    For i = 0 To rs.Fields.Count - 1
        mFldList = mFldList + rs.Fields(i).Name + "/"
    Next
    mFldList = "/" + Left(mFldList, Len(Trim(mFldList)) - 1) + "/"
    
    For Each ctlTemp In Screen.ActiveForm.Controls
        strFldname = Trim(LCase(Mid(ctlTemp.Name, 4, Len(ctlTemp.Name))))
        If IsInstr(strFldname, mFldList) Then
            Select Case LCase(TypeName(ctlTemp))
                Case LCase("CtxtBox"), LCase("ItxtBox"), LCase("NtxtBox"), LCase("GujtxtBox")
                    ctlTemp.Text = IfNullThen(rs.Fields(strFldname).Value, "")
                    
                Case LCase("Mskdat")
                    ctlTemp.Text = IfNullThen(rs.Fields(strFldname).Value, "__/__/____")
                    
                Case LCase("HlpNCode")
                    ctlTemp.CodeText = IfNullThen(rs.Fields(strFldname).Value, "")
                    'ctlTemp.GetNameText
                
                Case LCase("ComboBox")
                    ctlTemp.Text = IfNullThen(rs.Fields(strFldname).Value, "")
                    
                Case LCase("DTPicker")
                    ctlTemp.Value = rs.Fields(strFldname).Value
                
                Case LCase("CheckBox")
                    ctlTemp.Value = IIf(rs.Fields(strFldname).Value = "True", 1, 0)
                
                Case LCase("Shape"), LCase("Adodc"), LCase("Frame"), LCase("OptionButton"), LCase("SSTab"), LCase("ListView"), LCase("Label"), LCase("CommandButton")
                Case LCase("MSHFlexGrid")
            End Select
        End If
        'Debug.Print ctlTemp.Name
    Next
Exit Sub

errhndl:
    Resume Next
End Sub

Public Function IfNullThen(f_Value As Variant, f_RetValue As Variant) As Variant
    If IsNull(f_Value) Then
        IfNullThen = f_RetValue
    Else
        IfNullThen = f_Value
    End If
End Function

Public Function IsInstr(Word As String, str As String, Optional Seprator As String = "/") As Boolean
    Dim mChkWord As String
    
    str = IIf(Right(Trim(str), 1) = "/", str, Trim(str) + "/")
    
    mChkWord = Trim(Seprator) + Trim(Word) + Trim(Seprator)
    IsInstr = UCase(str) Like UCase("*" + mChkWord + "*")

End Function

Public Function SqlGenIns(ByVal f_TableName As String) As String
    Dim rsttemp As ADODB.Recordset
    Dim mFldList As String
    Dim mValueList As String
    
'---FiledList
    SQL = "select [name]"
    SQL = SQL & " from syscolumns where id in"
    SQL = SQL & " (select id from sysobjects where [name] = " & AQ(f_TableName) & ")"
    OpenAdoRst rsttemp, SQL
    
    With rsttemp
        If .RecordCount > 0 Then
            mFldList = .GetString(adClipString, , , ",")
        End If
    End With
    'If rsttemp.State = adStateOpen Then rsttemp.Close
    
'---ValueList
    SQL = "Select (case"
    SQL = SQL & "    when xtype in (56,108) then 'Val(Txt'"     'int,numeric
    SQL = SQL & "    when xtype in (167,175) then 'AQ(Txt'"     'varchar,char
    SQL = SQL & "    when xtype in (61) then 'ConvDatSql(Msk'  End)+[name]+'.Text)' as TypeName"
    SQL = SQL & " from syscolumns where id in"
    SQL = SQL & " (select id from sysobjects where [name] = " & AQ(f_TableName) & ")"
    OpenAdoRst rsttemp, SQL
    
    With rsttemp
        If .RecordCount > 0 Then
            mValueList = .GetString(adClipString, , , ",")
        End If
    End With
    'If rsttemp.State = adStateOpen Then rsttemp.Close
    
    SqlGenIns = "Insert into " & f_TableName & " (" & Left(mFldList, Len(mFldList) - 1) & ") Values (" & Left(mValueList, Len(mValueList) - 1) & ")"
    
    If rsttemp.State = adStateOpen Then rsttemp.Close
    Set rsttemp = Nothing
    
End Function

Public Sub ErrMsg()

Dim mFileName As String
Dim mFileNo As Integer
Dim mErrMsg As String


    mFileName = Trim$(gPathReport) & "ErrLog.Log"
    mFileNo = FreeFile
    
        Open mFileName For Output As mFileNo
    
        mErrMsg = Date & "-" & Time & _
                  " ErrNo:" & Err.Number & _
                  " ErrDesc:" & Err.Description & _
                  " ErrSource:" & Err.Source
                  
    
        Print #mFileNo, mErrMsg
    
        Close #mFileNo

    MsgBox "Error in Entry...!!!" & vbCrLf _
          & "Error Number : " & Err.Number & vbCrLf _
          & "Error Description : " & Err.Description, vbCritical
          
End Sub

Public Function SetMaxLength(f_Table As String, f_Field As String) As Integer
On Error GoTo errhndl
    
    Dim rs As ADODB.Recordset
    
    SQL = "select [name],isnull(Length,0) as Length from SysColumns "
    SQL = SQL & " Where id = Object_id(" & AQ(f_Table) & ")"
    SQL = SQL & " And Xtype in (167,175)" '167 Varchar,175 Char
    SQL = SQL & " And [Name]  = " & AQ(f_Field)
    OpenAdoRst rs, SQL
    If rs.RecordCount > 0 Then
        SetMaxLength = rs.Fields("legnth").Value
    Else
        SetMaxLength = 0
    End If
    CloseAdoRst rs
    
Exit Function
    
errhndl:
    Resume Next
    
End Function

Public Function AdoIsField(f_FldName As String, f_Rst As ADODB.Recordset) As Boolean
Dim mRetVal As Boolean
Dim i As Integer

For i = 0 To f_Rst.Fields.Count - 1
    If LCase(f_FldName) = LCase(f_Rst.Fields(i).Name) Then
        mRetVal = True
        Exit For
    Else
        mRetVal = False
    End If
Next

AdoIsField = mRetVal

End Function

Public Function adofieldDet(s_Table As String, s_fieldname As String, Optional s_Return As String, Optional s_Cnn As ADODB.Connection) As Variant
Dim rstFlddet As ADODB.Recordset
Dim mRetVal As Variant
    
    If s_Cnn Is Nothing Then Set s_Cnn = gCnnMst
    
    SQL = "select c.name fld_name,t.name fld_type,c.length fld_len,c.xprec fld_prec,c.scale fld_scale "
    SQL = SQL & " from syscolumns c,sysobjects o,systypes t"
    SQL = SQL & " where c.id=o.id and o.name='" & s_Table & "' and t.xtype=c.xtype and c.name='" & s_fieldname & "'"
    OpenAdoRst rstFlddet, SQL, , , , s_Cnn
    
    With rstFlddet
    If .RecordCount > 0 Then
        If Len(s_Return) <> 0 Then
            Select Case LCase(s_Return)
                Case "type"
                    mRetVal = .Fields("fld_type").Value
                Case "length"
                    mRetVal = .Fields("fld_len").Value
                Case "prec"
                    mRetVal = .Fields("fld_prec").Value
                Case "scale"
                    mRetVal = .Fields("fld_scale").Value
            End Select
        Else
            mRetVal = True
        End If
    Else
        mRetVal = False
    End If
    End With
    adofieldDet = mRetVal
    rstFlddet.Close
    Set rstFlddet = Nothing
End Function

Public Sub AlterTableColumn(s_Table As String, s_Column As String, s_Type As String, Optional s_Len As Integer, Optional s_dec As Integer, Optional s_AddorAlter As String = "add", Optional s_Constaint As String, Optional s_Cnn As ADODB.Connection)
Dim mConstraint As String
Dim mDfVal As Variant
    
    If s_Cnn Is Nothing Then Set s_Cnn = gCnnMst
    
    If Len(s_Constaint) = 0 Then
        mConstraint = ""
    Else
        Select Case Left(LCase(s_Constaint), 8)
            Case "default("
                mConstraint = " constraint DF__" & s_Table & "__" & s_Column & " " & s_Constaint
        End Select
    End If
    
    If Val(s_Len) = 0 And Val(s_dec) = 0 Then  '---date type
        If s_AddorAlter = "add" Then
            SQL = "alter table " & s_Table & " " & s_AddorAlter & " " & s_Column & " " & s_Type
        Else
            If LCase(s_Constaint) = "not null" Then
                SQL = "alter table " & s_Table & " " & s_AddorAlter & " column " & s_Column & " " & s_Type & " " & s_Constaint
            Else
                SQL = "alter table " & s_Table & " " & s_AddorAlter & " column " & s_Column & " " & s_Type
            End If
        End If
    ElseIf Val(s_dec) = 0 Then   '---char type
        If s_AddorAlter = "add" Then
            SQL = "alter table " & s_Table & " " & s_AddorAlter & " " & s_Column & " " & s_Type & "(" & s_Len & ") " & mConstraint
        Else
            If LCase(s_Constaint) = "not null" Then
                SQL = "alter table " & s_Table & " " & s_AddorAlter & " column " & s_Column & " " & s_Type & "(" & s_Len & ") " & s_Constaint
            Else
                SQL = "alter table " & s_Table & " " & s_AddorAlter & " column " & s_Column & " " & s_Type & "(" & s_Len & ")"
            End If
        End If
    Else  '---numeric type
        If s_AddorAlter = "add" Then
            SQL = "alter table " & s_Table & " " & s_AddorAlter & " " & s_Column & " " & s_Type & "(" & s_Len & "," & s_dec & ") " & mConstraint
        Else
            SQL = "alter table " & s_Table & " " & s_AddorAlter & " column " & s_Column & " " & s_Type & "(" & s_Len & "," & s_dec & ")"
        End If
    End If
    
    s_Cnn.Execute SQL
End Sub

