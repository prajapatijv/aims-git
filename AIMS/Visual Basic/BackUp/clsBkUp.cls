VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBkUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BackUpDb(s_Db As String, s_BackUpPath As String, s_Cnn As ADODB.Connection, Optional IsMsgPop As Boolean = True)
On Error GoTo errhndl
MP vbHourglass

    Dim SQL As String
    
'    If AdoIsDatabase(s_Db, s_Cnn) = False Then
'        MP vbDefault
'        Exit Sub
'    End If
    
    Dim mBackUpPath As String
    If Len(s_BackUpPath) <= 0 Then
        s_BackUpPath = App.Path
    End If
    
    mBackUpPath = s_BackUpPath & Trim(s_Db) & "_" & Format(Now, "YYMMDD") & "_" & Format(Now, "HHMM")
        
    If IsFile(mBackUpPath) Then
        Kill mBackUpPath
    End If
        
    SQL = "BackUp Database " & s_Db & " To Disk = " & AQ(mBackUpPath & ".dat")
    s_Cnn.Execute SQL
    
    If IsMsgPop = True Then
        MsgBox "BackUp Completed Successfully", vbInformation
    End If

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Function AdoIsDatabase(s_Db As String, s_Cnn As ADODB.Connection) As Boolean
Dim rst As ADODB.Recordset
Dim mRtnVal As Boolean

    Dim SQL As String

    SQL = "select name from master.dbo.sysDatabases where name='" & s_Db & "'"
    
    Set rst = New ADODB.Recordset
    
    OpenAdoRst rst, SQL, , , , s_Cnn
    
    If rst.RecordCount > 0 Then
        mRtnVal = True
    Else
        mRtnVal = False
    End If
    
    rst.Close
    Set rst = Nothing
    AdoIsDatabase = mRtnVal

End Function

Private Sub MP(Pointer As MousePointerConstants)
    Screen.MousePointer = Pointer
End Sub

Private Function IsFile(f_filename As String) As Boolean

If InStr(f_filename, ".") = 0 Then
    If Len(Trim(Dir(f_filename, vbDirectory))) = 0 Then
        IsFile = False
    Else
        IsFile = True
    End If
Else
    If Len(Trim(Dir(f_filename))) = 0 Then
        IsFile = False
    Else
        IsFile = True
    End If
End If

End Function

Private Function AQ(f_str As String, Optional f_QuoteChar As String = "'", Optional f_Addcoma As Boolean = False) As String
    AQ = f_QuoteChar + f_str + f_QuoteChar
    If f_Addcoma Then AQ = AQ + ","
End Function

Private Sub ErrMsg()

Dim mFileName As String
Dim mFileNo As Integer
Dim mErrMsg As String


    mFileName = Trim$(App.Path) & "\ErrLog.Log"
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


Public Sub OpenAdoRst(s_AdoRst As ADODB.Recordset, s_Source, Optional s_CursorType As ADODB.CursorTypeEnum = adOpenDynamic, Optional s_CmdType As CommandTypeEnum = adCmdText, Optional s_LckType As LockTypeEnum = adLockOptimistic, Optional s_Cnn As ADODB.Connection)

    Set s_AdoRst = New ADODB.Recordset
    s_AdoRst.Open s_Source, s_Cnn, s_CursorType, s_LckType
    
End Sub

