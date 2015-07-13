Attribute VB_Name = "sqlProc"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B8300C701CF"
Option Explicit

Public Sub SetDBOptions()
    gCnnMst.Execute "sp_dboption '" & gMdbMst & "', 'select into/bulkcopy', 'true'"
End Sub

Public Sub BackUpDb(s_Db As String, Optional IsMsgPop As Boolean = True)
On Error GoTo errhndl
MP vbHourglass
    
    If AdoIsDatabase(s_Db) = False Then
        MP vbDefault
        Exit Sub
    End If
    
    Dim mBackUpPath As String
    If Len(gBackUpPath) <= 0 Then
        gBackUpPath = App.Path
    End If
    
    mBackUpPath = gBackUpPath & Trim(s_Db) & "_" & Format(Now, "YYMMDD") & "_" & Format(Now, "HHMM")
        
    If IsFile(mBackUpPath) Then
        Kill mBackUpPath
    End If
        
    SQL = "BackUp Database " & s_Db & " To Disk = " & AQ(mBackUpPath)
    gCnnMst.Execute SQL
    
    If IsMsgPop = True Then
        MsgBox "BackUp Completed Successfully", vbInformation
    End If

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub
