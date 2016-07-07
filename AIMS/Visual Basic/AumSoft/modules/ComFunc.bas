Attribute VB_Name = "ComFunc"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B8300860316"
Option Explicit

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'Public Const GWL_EXSTYLE = (-20)
'Public Const WS_EX_LAYERED = &H80000
'Public Const LWA_ALPHA = &H2&

Dim fso As New FileSystemObject

Public Sub TransperentForm(s_Form As Form, s_Opacity As Integer)
'    If s_Opacity < 0 Then s_Opacity = 150
    
'    Call SetWindowLong(s_Form.hwnd, GWL_EXSTYLE, GetWindowLong(s_Form.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
'    Call SetLayeredWindowAttributes(s_Form.hwnd, 0, s_Opacity, LWA_ALPHA)
End Sub

Public Function Readfromini(Read_var As String, Seq_file As String)
Dim input_file As String, chk_var As String
Dim Free_area As Integer

On Error GoTo IniReadErr
    
    Readfromini = ""
        
    chk_var = Read_var
    
    If InStr(Seq_file, ".") = 0 Then
        input_file = Seq_file + ".ini"
    Else
        input_file = Seq_file
    End If
    
    Free_area = FreeFile
    
    'If Not IsFileMsg(input_file) Then Readfromini = "": Exit Function
    
    Open input_file For Input As Free_area
    Do While Not EOF(Free_area)
        Line Input #Free_area, chk_var
        If Len(chk_var) <> 0 Then
            If Trim(UCase(Mid(chk_var, 1, InStr(chk_var, "=") - 1))) = UCase(Read_var) Then
                Readfromini = Trim(Mid(chk_var, InStr(chk_var, "=") + 1, 200))
                Exit Do
            Else
                Readfromini = ""
            End If
        End If
    Loop
    Close Free_area
    
    
    Exit Function
    
IniReadErr:
    Exit Function
End Function
Public Function Writetoini(Txt_string As String, Seq_file As String)
Dim mFound As Boolean
Dim Free_area, i, mTot_line As Integer
Dim chk_var, input_file, Read_var, mIni_line(20) As String

    If InStr(Seq_file, ".") = 0 Then
        input_file = Seq_file + ".ini"
    End If
    
    Free_area = FreeFile

    Open Seq_file For Input As Free_area
    Do While Not EOF(Free_area)
        Input #Free_area, chk_var
        If Mid(chk_var, 1, InStr(chk_var, "=") - 1) = Mid(Txt_string, 1, InStr(Txt_string, "=") - 1) Then
            mFound = True
            Exit Do
        Else
            mFound = False
        End If
    Loop
    Close Free_area
    
    Select Case mFound
    Case True
         Open Seq_file For Input As Free_area
         i = 0
         Do While Not EOF(Free_area)
             i = i + 1
             Input #Free_area, mIni_line(i)
             If Mid(mIni_line(i), 1, InStr(mIni_line(i), "=")) = Mid(Txt_string, 1, InStr(Txt_string, "=")) Then
                 mIni_line(i) = Txt_string
             End If
             mTot_line = i
         Loop
         Close Free_area
    
         Kill Seq_file
         Open Seq_file For Append As Free_area
         For i = 1 To mTot_line
             Print #Free_area, mIni_line(i)
         Next
         Close Free_area
     
     Case False
         Open Seq_file For Append As Free_area
             Print #Free_area, Txt_string
         Close Free_area
     
     End Select

End Function

Public Function ChartoAsc(f_Text As String) As String
Dim mAsc_str As String
Dim i As Integer
    
    If IsNumeric(f_Text) Then str (f_Text)
    
    mAsc_str = ""

    For i = 1 To Len(f_Text)
        mAsc_str = mAsc_str + Chr(Asc(Mid(f_Text, i, 1)) + (90 + i * 2))
    Next
    ChartoAsc = mAsc_str

End Function

Public Function AsctoChar(f_Text As String) As String
Dim mChar_str As String
Dim i As Integer

On Error GoTo Lc

    If IsNull(f_Text) Then
        AsctoChar = ""
        Exit Function
    End If
    
    mChar_str = ""
    For i = 1 To Len(f_Text)
        mChar_str = mChar_str + Chr(Asc(Mid(f_Text, i, 1)) - (90 + i * 2))
    Next
    AsctoChar = mChar_str

Exit Function
Lc:
    'Errmsg pErr_Loca & "asctochar"
    Resume Next
End Function

Public Function CountSign(line As String, Optional Sign As String = "/") As Integer
Dim mCount As Integer
Dim i As Integer

mCount = 0
For i = 1 To Len(Trim(line))
    If Mid(line, i, 1) = Sign Then mCount = mCount + 1
Next

CountSign = mCount

End Function
Public Function RandomNo() As Double
    Randomize
    RandomNo = Int(100000 * Rnd)
End Function

Public Function CheckDateFormat() As Boolean
    
If Format("10/12/2000", "dd/mm/yyyy") <> "10/12/2000" Or _
    Right((FormatDateTime("10/12/2000", vbShortDate)), 4) <> "2000" Then
    MsgBox "System date setting is not in dd/mm/yyyy format." & vbCrLf & "Please change the format from {Control Panel - Regional Setting - Date} & then Continue.", vbOKOnly
    CheckDateFormat = False
Else
    CheckDateFormat = True
End If

End Function

Public Function AddQuote(s_Str As String, Optional s_QuoteChar As String = "'") As String
    AddQuote = s_QuoteChar + s_Str + s_QuoteChar
End Function

Public Function AQ(f_str As String, Optional f_QuoteChar As String = "'", Optional f_Addcoma As Boolean = False) As String
    AQ = f_QuoteChar + f_str + f_QuoteChar
    If f_Addcoma Then AQ = AQ + ","
End Function

Public Function GetComp_Id() As Integer
    GetComp_Id = Val(Environ("Comp_id"))
End Function

Public Function ConvDatSql(s_text As String, Optional s_backend As BackEnd_Type) As String

    ConvDatSql = "'" & Format(s_text, "mm/dd/yyyy") & "'"

End Function
Public Sub CreateIniFile()
Dim mIniFile As String
Dim mFileNo As Integer

On Error GoTo Lc

    mIniFile = Trim(App.Path) & "\" & SetPkgName() + ".ini"
    
    If Not IsFile(mIniFile) Then
        mFileNo = FreeFile
        
        Open mIniFile For Output As mFileNo
                Print #mFileNo, "servername="""""
                Print #mFileNo, "invdb=" + gPkgName + "01"
                Print #mFileNo, "ReportPath=" + App.Path + "\reports\"
                Print #mFileNo, "WorkingPath=" + App.Path + "\"
                Print #mFileNo, "ResourcePath=" + App.Path + "\images\"
                Print #mFileNo, "BackUpPath=" + App.Path + "\images\"
                Print #mFileNo, "be=2"
        Close #mFileNo
        
        MsgBox UCase(mIniFile) & " is been Created for path settings" & vbCrLf & vbCrLf, vbOKOnly
    End If
    
Exit Sub
Lc:
    Resume Next
End Sub

Public Sub SetFocusTo(s_Object As Object)
    If s_Object.Visible Then If s_Object.Enabled Then s_Object.SetFocus
End Sub

Public Sub CenterFrmChild(s_Form As Form)
With s_Form
  .Top = (mdiMainMenu.ScaleHeight / 2) - (.Height / 2)
  .Left = (mdiMainMenu.ScaleWidth / 2) - (.Width / 2)
End With
End Sub
Public Sub CenterFrmNonChild(s_Form As Form)
With s_Form
  .Top = (Screen.Height / 2) - (.Height / 2)
  .Left = (Screen.Width / 2) - (.Width / 2)
End With
End Sub

Public Sub FillCombo(s_Combo As ComboBox, s_List As String, Optional s_Clear As Boolean = True, Optional s_Sign As String = "~")
Dim i As Integer
Dim mAddItem As String
Dim mFullLine As String

With s_Combo
mFullLine = s_List

If s_Clear Then .Clear
i = 0

For i = 0 To CountSign(s_List, s_Sign) - 1
    mAddItem = Left(mFullLine, InStr(mFullLine, s_Sign) - 1)
    mFullLine = LChop(mFullLine, Len(mAddItem) + 2)
    .AddItem mAddItem
Next i

'If CountSign(s_List, s_Sign) - 1 = 0 Then .AddItem mFullLine
If Len(mFullLine) > 0 Then .AddItem mFullLine

If .ListCount > 0 Then .ListIndex = 0

End With

End Sub

Public Function IsFile(f_filename As String) As Boolean

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
Public Function IsFileMsg(f_filename As String, Optional f_AddMsg As String) As Boolean

If IsFile(f_filename) Then
    IsFileMsg = True
Else
    MsgBox "Hey FILE MISSING - " & UCase(f_filename), vbCritical, "Error : File Missing"
    IsFileMsg = False
End If

End Function
Public Function LChop(s_String As String, s_CharLen) As String

LChop = Mid(s_String, s_CharLen, Len(s_String) - Len(s_CharLen))

End Function


Public Sub ClearScreen()
Dim ctlTemp As Control
    For Each ctlTemp In Screen.ActiveForm.Controls
        'Debug.Print LCase(TypeName(ctlTemp))
        Select Case LCase(TypeName(ctlTemp))
            Case "textbox", "datacombo", "ctxtbox", "itxtbox", "ntxtbox", LCase("GujtxtBox")
                ctlTemp.Text = ""
            Case "combobox"
                If ctlTemp.Style <> 2 Then ctlTemp.Text = ""
            Case "listview"
                ctlTemp.ListItems.Clear
            Case "mshflexgrid", "msflexgrid"
                ctlTemp.Clear
            Case "maskedbox", "mskdat"
                ctlTemp.Text = "__/__/____"
            Case "hlpncode"
                ctlTemp.CodeText = 0
                ctlTemp.NameText = ""
        End Select
    Next
End Sub

Public Sub CreateMdb(s_MdbName As String)
Dim WsCreate As Workspace
Dim DbCreate As Database
Dim mPW As String

' vbHourglass

    If IsFile(s_MdbName) Then Exit Sub
    
    If Len(Trim(Dir(s_MdbName))) = 0 Then
        Set WsCreate = DBEngine.Workspaces(0)
        Set DbCreate = WsCreate.CreateDatabase(s_MdbName, dbLangGeneral, dbEncrypt)
    End If
    
    'DbCreate.NewPassword "", gMdbPwd
    
    Set DbCreate = Nothing
    Set WsCreate = Nothing
        
End Sub

Public Sub AgentMsg(s_String As String)
'''    mdiMainMenu.Agent1.Characters.Load "x", "D:\JNB\prps\" & Left(Trim(WeekdayName(Weekday(Date))), 3) & ".prp"
'''    'mdiMainMenu.Agent1.Characters("x").Left = 1000
'''    'mdiMainMenu.Agent1.Characters("x").Top = 2000
'''    mdiMainMenu.Agent1.Characters("x").Show
'''    mdiMainMenu.Agent1.Characters("x").Speak s_String
'''
'''    mdiMainMenu.Agent1.Characters("x").Hide (20)
'''    'mdiMainMenu.Agent1.Characters.Unload ("x")
End Sub

Public Sub VisibleNoVisibleBtn(s_Visible As Boolean, Optional s_FormUnload As Boolean = False)
Dim i As Integer

    With mdiMainMenu
        .TbrMain.ButtonHeight = 350
        .TbrMain.ButtonWidth = 350
        For i = 1 To .TbrMain.Buttons.Count
           .TbrMain.Buttons(i).Visible = s_Visible
           .TbrMain.Buttons(i).Enabled = s_Visible
        Next
        
        If Not s_FormUnload Then
            SetUserRights Screen.ActiveForm.Name, gUser, False
        End If
        
        .TbrMain.Buttons(btnsave).Enabled = False
        .TbrMain.Buttons(btnSaveNAdd).Enabled = False
        .TbrMain.Buttons(btnCancel).Enabled = False
                
    End With

End Sub

Public Sub GrabActiveControl()
    On Error Resume Next
        
    If Screen.ActiveForm.Controls.Count > 0 Then
        If (Screen.ActiveForm.ActiveControl Is Nothing) Then Exit Sub
        Set Screen.ActiveForm.mActCtrl = Screen.ActiveForm.ActiveControl
    End If
    
    'Screen.ActiveForm.KeyPreview = True
End Sub

Public Sub SetActiveModeNControl(s_EntryMode As String)
    Select Case LCase(s_EntryMode)
        Case "add"
            Screen.ActiveForm.EnableDisable True
            BtnPressed mdiMainMenu.TbrMain.Buttons(btnadd)
            
            SetFocusTo Screen.ActiveForm.mActCtrl
        Case "edit"
            Screen.ActiveForm.EnableDisable True
            BtnPressed mdiMainMenu.TbrMain.Buttons(btnedit)
            SetFocusTo Screen.ActiveForm.mActCtrl
        Case "view"
            Screen.ActiveForm.EnableDisable True
            BtnPressed mdiMainMenu.TbrMain.Buttons(btnview)
            SetFocusTo Screen.ActiveForm.mActCtrl
        Case Else
            Screen.ActiveForm.EnableDisable False
    End Select
End Sub

Public Sub FormKeyDown(s_KeyCode As Integer, s_Shift As Integer, Optional s_ApplyEsc As Boolean = True)
    'MsgBox s_KeyCode
    Select Case s_KeyCode
        Case vbKeyReturn
            If LCase(TypeName(Screen.ActiveForm.ActiveControl)) <> "mshflexgrid" And _
               LCase(TypeName(Screen.ActiveForm.ActiveControl)) <> "commandbutton" Then
                SendKeys "{TAB}"
                s_KeyCode = 0
            End If
            
        Case vbKeyEscape
            If Screen.ActiveForm.ActiveControl.TabIndex = 0 Then
                Unload Screen.ActiveForm
                s_KeyCode = 0
            ElseIf LCase(TypeName(Screen.ActiveForm.ActiveControl)) <> "mshflexgrid" And _
               LCase(TypeName(Screen.ActiveForm.ActiveControl)) <> "commandbutton" Then
                If s_ApplyEsc Then SendKeys "+{TAB}"
                DoEvents: DoEvents
                DoEvents: DoEvents
                's_KeyCode = 0
            End If
            
        
        Case vbKeyPageUp
            If mdiMainMenu.TbrMain.Buttons(btnprev).Enabled = True Then
                mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnprev)
                s_KeyCode = 0
            End If
        Case vbKeyPageDown
            If mdiMainMenu.TbrMain.Buttons(btnnext).Enabled = True Then
                mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnnext)
                s_KeyCode = 0
            End If
        Case vbKeyHome
            If mdiMainMenu.TbrMain.Buttons(btnfirst).Enabled = True Then
                mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnfirst)
                s_KeyCode = 0
            End If
        Case vbKeyEnd
            If mdiMainMenu.TbrMain.Buttons(btnlast).Enabled = True Then
                mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnlast)
                s_KeyCode = 0
            End If
        Case vbKeyF1
            If LCase(TypeName(Screen.ActiveForm.ActiveControl)) = LCase("HlpNCode") Then
                F1Help Screen.ActiveForm.ActiveControl.TableName
            End If
    End Select
    
    If s_Shift = 4 And s_KeyCode = vbKeyA Then 'Add
        If mdiMainMenu.TbrMain.Buttons(btnadd).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnadd)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyE Then 'Edit
        If mdiMainMenu.TbrMain.Buttons(btnedit).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnedit)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyD Then 'Del
        If mdiMainMenu.TbrMain.Buttons(btndel).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btndel)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyS Then 'Save
        If mdiMainMenu.TbrMain.Buttons(btnsave).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnsave)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyC Then 'Cancel
        If mdiMainMenu.TbrMain.Buttons(btnCancel).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnCancel)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyP Then ' Print
        If mdiMainMenu.TbrMain.Buttons(btnprint).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnprint)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyV Then ' View
        If mdiMainMenu.TbrMain.Buttons(btnview).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnview)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyX Then ' Exit
        If mdiMainMenu.TbrMain.Buttons(btnExit).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnExit)
            s_KeyCode = 0
        End If
    ElseIf s_Shift = 4 And s_KeyCode = vbKeyF Then ' Find
        If mdiMainMenu.TbrMain.Buttons(btnFind).Enabled = True Then
            mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnFind)
            s_KeyCode = 0
        End If
    End If
    
End Sub

Public Sub BtnPressed(ByVal Button As MSComctlLib.Button)
    Dim i As Integer
    
    
    With mdiMainMenu
        
        For i = 1 To .TbrMain.Buttons.Count
            .TbrMain.Buttons(i).Enabled = False
        Next
        .TbrMain.Buttons(btnsave).Enabled = False
        .TbrMain.Buttons(btnSaveNAdd).Enabled = False
        .TbrMain.Buttons(btnCancel).Enabled = False
        
        Select Case LCase(Button.Tag)
            Case "add"
                .TbrMain.Buttons(btnsave).Enabled = True
                .TbrMain.Buttons(btnSaveNAdd).Enabled = True
                .TbrMain.Buttons(btnCancel).Enabled = True
                
            Case "edit"
                .TbrMain.Buttons(btnsave).Enabled = True
                .TbrMain.Buttons(btnSaveNAdd).Enabled = True
                .TbrMain.Buttons(btnCancel).Enabled = True
            
            Case "del"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
                
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "save"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
            
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "saveaddnew" 'save and add
                .TbrMain.Buttons(btnsave).Enabled = True
                .TbrMain.Buttons(btnSaveNAdd).Enabled = True
                .TbrMain.Buttons(btnCancel).Enabled = True
            
            Case "cancel"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
            
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "print"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
                
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "view"
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = True
                .TbrMain.Buttons(btnadd).Enabled = True
                .TbrMain.Buttons(btnedit).Enabled = True
                
            Case "first"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
            
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "next"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
            
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "prev"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
            
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
            
            Case "last"
                For i = 1 To .TbrMain.Buttons.Count
                    .TbrMain.Buttons(i).Enabled = True
                Next
            
                .TbrMain.Buttons(btnsave).Enabled = False
                .TbrMain.Buttons(btnSaveNAdd).Enabled = False
                .TbrMain.Buttons(btnCancel).Enabled = False
                            
            'Case "exit"
            '    Screen.ActiveForm.EntryExit
            
            'Case "quit"
            '    EntryQuit
        End Select
    End With
End Sub

Public Function ValidateControl() As Boolean
Dim mCtrl As Control
Dim mStr As String
Dim strFldname As String

    ValidateControl = True
    
    For Each mCtrl In Screen.ActiveForm.Controls
        If LCase(TypeName(mCtrl)) = "commondialog" Then
            'Do nothing
        Else
            If LCase(TypeName(mCtrl)) = "label" And mCtrl.Visible Then
                If InStr(1, mCtrl.Caption, "*", vbTextCompare) > 0 Then
                      mStr = Mid(mCtrl.Name, 4) & "/" & mStr
                End If
            End If
        End If
    Next
        
    For Each mCtrl In Screen.ActiveForm.Controls
        strFldname = Trim(LCase(Mid(mCtrl.Name, 4, Len(mCtrl.Name))))
        If IsInstr(strFldname, mStr) Then
            Select Case LCase(TypeName(mCtrl))
                Case LCase("CtxtBox"), LCase("TextBox"), LCase("GujtxtBox")
                    If Len(mCtrl.Text) <= 0 Then
                        ValidateControl = False
                        SetFocusTo mCtrl
                        Exit For
                    End If
                Case LCase("ItxtBox"), LCase("NtxtBox")
                    If Val(mCtrl.Text) <= 0 Then
                        ValidateControl = False
                        SetFocusTo mCtrl
                        Exit For
                    End If
                Case LCase("HlpNCode")
                    If Val(mCtrl.CodeText) <= 0 Then
                        ValidateControl = False
                        SetFocusTo mCtrl
                        Exit For
                    End If
                Case LCase("Mskdat")
                    If Not IsDate(mCtrl.Text) Then
                        ValidateControl = False
                        SetFocusTo mCtrl
                        Exit For
                    End If
                Case LCase("ComboBox")
                    If Len(mCtrl.Text) <= 0 Then
                        ValidateControl = False
                        SetFocusTo mCtrl
                        Exit For
                    End If
            
            End Select
        End If
    Next
End Function
Public Sub OpenRs()
Dim rs As New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
    Select Case Screen.ActiveForm.Tag
        Case "compmst"
            rs.Open "SELECT CODE,NAME FROM CompMast", gCnnMst, adOpenDynamic, adLockOptimistic
     End Select
     'Call frmHelpName.fillgrid(rs)
End Sub

Public Sub MP(Pointer As MousePointerConstants)
    Screen.MousePointer = Pointer
End Sub

Public Sub AskSave(s_EntryNo As String, s_FocusTo As Control, s_EntryMode As String)
    
    If LCase(s_EntryMode) = "view" Then Exit Sub
    If MsgBox("Want to Save Entry No : " & s_EntryNo & "..???.. ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        'Screen.ActiveForm.EntrySave
        mdiMainMenu.tbrMain_ButtonClick mdiMainMenu.TbrMain.Buttons(btnsave)
    Else
        SetFocusTo s_FocusTo
    End If

End Sub

Public Sub OnlineNums(s_KeyAscii, s_Control As Control, Optional s_Decimals As Integer = 2)
    Select Case s_KeyAscii
        Case vbKey0 To vbKey9, 46, vbKeyBack
            If InStr(1, s_Control.Text, ".", vbTextCompare) > 0 Then
                If s_KeyAscii <> 46 Then
                    If s_Control.MaxLength = 0 Then
                        s_Control.MaxLength = Len(s_Control.Text) + s_Decimals
                    End If
                Else
                    s_KeyAscii = 0
                End If
            Else
                s_Control.MaxLength = 0
            End If
        Case Else
            s_KeyAscii = 0
    End Select
End Sub

Public Sub OnlineCaps(s_Control As Control)
        
    Dim i As Integer
    i = s_Control.SelStart
    s_Control.Text = UCase(s_Control.Text)
    s_Control.SelStart = i

End Sub


Public Sub ControlGotFocus(s_Control As Control)
    s_Control.SelStart = 0
    s_Control.SelLength = Len(s_Control.Text)
End Sub

Public Function DtaTime() As String
    DtaTime = Replace(Time, " ", "")
End Function

Public Sub SetFlexFixedColCheckBoxes(s_Msf As MSHFlexGrid, s_CheckColumn As Integer, Optional s_CheckAll As Boolean = True, Optional s_StartRowNo As Integer = 1)
    Dim i As Integer
    
    With s_Msf
        For i = s_StartRowNo To .Rows - 1
            .Row = i: .Col = s_CheckColumn
            .CellFontName = "Marlett"
            If s_CheckAll Then
                .TextMatrix(i, s_CheckColumn) = "a"
            Else
                .TextMatrix(i, s_CheckColumn) = "r"
            End If
        Next
    End With
    
End Sub

Public Sub SetGridColGujFont(s_Msf As MSHFlexGrid, s_Column As Integer, Optional s_iFontSize As Integer = 12, Optional s_StartingRow As Integer = 0)
    Dim i As Integer
    With s_Msf
        For i = s_StartingRow To .Rows - 1
            .Row = i: .Col = s_Column
            .CellFontName = gGujaratiFontName
            .CellFontSize = s_iFontSize
        Next
    End With
End Sub

Public Sub SetGridRowColor(s_Msh As MSHFlexGrid, s_RowNum As Integer, s_Color As OLE_COLOR, Optional s_FromCol As Integer = 0, Optional s_FontBold As Boolean = False)

    Dim iCol As Integer
    
    If s_RowNum = -1 Then Exit Sub
    
    With s_Msh
        .Row = s_RowNum
        For iCol = s_FromCol To .Cols - 1
            .Col = iCol
            .CellBackColor = s_Color
            .CellFontBold = s_FontBold
        Next
    End With
    
End Sub


Public Sub SetUserRights(s_Form As String, s_Uid As String, s_ApplySys As Boolean)

    If s_ApplySys Then
        Dim rst As ADODB.Recordset
        
        SQL = "Select u.*,f.FormName as FormName,f.FormGrp as FormGrp "
        SQL = SQL & " From UserRights u "
        SQL = SQL & " Inner Join FormCollection f "
        SQL = SQL & " On u.FormId=f.FormId "
        SQL = SQL & " Where f.FormName = " & AQ(s_Form)
        SQL = SQL & " And u.uid = " & Val(s_Uid)
        
        OpenAdoRst rst, SQL, , , , gCnnMst
        
        With rst
            If .RecordCount > 0 Then
                If LCase(.Fields("FormGrp").Value) = "r" Then 'Report
                    If Not (CBool(IfNullThen(.Fields("report").Value, 0))) Then
                        Unload Screen.ActiveForm
                    End If
                Else
                    gAdd = CBool(IfNullThen(.Fields("add").Value, 0))
                    gEdit = CBool(IfNullThen(.Fields("edit").Value, 0))
                    gDel = CBool(IfNullThen(.Fields("del").Value, 0))
                    gPrint = CBool(IfNullThen(.Fields("print").Value, 0))
                    gNevigate = CBool(IfNullThen(.Fields("nevigate").Value, 0))
                End If
            Else
            
            End If
            
        End With
    Else
        gAdd = True
        gEdit = True
        gDel = True
        gPrint = True
        gNevigate = True
    End If
End Sub

Public Sub F1Help(s_MstTbl As String)

    Select Case LCase(s_MstTbl)
        Case "acmast"
            mdiMainMenu.mnuMasterArr1_Click (0)
        Case "patmast"
            mdiMainMenu.mnuMasterArr1_Click (1)
        Case "matmast"
            mdiMainMenu.mnuMasterArr1_Click (2)
        Case "palletemast"
            mdiMainMenu.mnuMasterArr1_Click (3)
        
        Case "gdnmast"
            mdiMainMenu.mnuMasterArr1_Click (5)
        Case "cbmast"
            mdiMainMenu.mnuMasterArr1_Click (6)
        Case "rackmast"
            mdiMainMenu.mnuMasterArr1_Click (7)
        Case "drwmast"
            mdiMainMenu.mnuMasterArr1_Click (8)
        Case "grpmast"
            mdiMainMenu.mnuMasterArr1_Click (9)
        Case "unitmast"
            mdiMainMenu.mnuMasterArr1_Click (10)
            
        Case "compmast"
        Case "rawmatmast"
            mdiMainMenu.mnuMasterArr1_Click (13)
        Case "costcentmast"
            mdiMainMenu.mnuMasterArr1_Click (14)
        Case "grademast"
            mdiMainMenu.mnuMasterArr1_Click (18)
    End Select
    
End Sub

Public Sub SetComboMaxLength(s_Control As ComboBox, s_Lengh As Integer)
    If Len(s_Control.Text) > s_Lengh Then
        s_Control = Left(s_Control.Text, s_Lengh)
        s_Control.SelStart = Len(s_Control.Text)
    End If
End Sub

Public Sub SetDefaConum(s_HlpCtrl As HlpNCode)
    If Val(s_HlpCtrl.CodeText) <= 0 Then
        s_HlpCtrl.CodeText = gDefaComp
    End If
End Sub

Public Sub CenterFormCaption(s_Form As Form, s_FormCaption As String)
    If s_Form.WindowState <> vbMinimized Then
        s_Form.Caption = Space((Val((s_Form.ScaleWidth) / 69) - Len(Trim(s_FormCaption))) / 2) & ":: " & Trim(s_FormCaption) & " ::"
    Else
        s_Form.Caption = ":: " & Trim(s_FormCaption) & " ::"
    End If
End Sub

Public Function CheckUniqueDocNo(f_TableName As String, f_Where As String) As Boolean
On Error GoTo errhndl
MP vbHourglass
    
    SQL = "Select 'True' "
    SQL = SQL & "From " & f_TableName
    SQL = SQL & " Where 1=1 "
    
    If Len(f_Where) > 0 Then
        SQL = SQL & " And " & f_Where
    End If
        
    If gCnnMst.Execute(SQL).RecordCount > 0 Then
        CheckUniqueDocNo = False
        MsgBox "Given Document Number Already Exists...!!!" & vbCrLf & vbCrLf & "Please Provide Another Document Number.", vbInformation
    Else
        CheckUniqueDocNo = True
    End If
   
MP vbDefault
Exit Function
errhndl:
    ErrMsg
    Resume Next
    
End Function

Public Function GetAcMastFlds(s_Alias As String) As String
    Dim mMstLst As String
    mMstLst = " ," & Trim(s_Alias) & ".[name] as PtyName" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[AlsName] as PtyAliasName" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[ContactPer] as PtyContacePer" & vbCrLf
    
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[add1] as PtyAdd1" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[add2] as PtyAdd2" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[add3] as PtyAdd3" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[City] as PtyCity" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Pncd] as PtyPncd" & vbCrLf

    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Dist] as PtyDist" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[State] as PtyState" & vbCrLf

    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Phone] as PtyPhone" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Phone1] as PtyPhone1" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Phone2] as PtyPhone2" & vbCrLf

    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Mobile] as PtyMobile" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Fax] as PtyFax" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[Email] as PtyEmail" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[WWW] as PtyWWW" & vbCrLf

    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[CstNo] as PtyCstNo" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[GstNo] as PtyGstNo" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[PanNo] as PtyPanNNo" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[EccNo] as PtyEccNo" & vbCrLf
    mMstLst = mMstLst & " ," & Trim(s_Alias) & ".[VatNo] as PtyVatNo" & vbCrLf
    
    GetAcMastFlds = mMstLst
    
End Function

Public Function TLen(f_Expr As Variant)
    TLen = Len(Trim(f_Expr))
End Function

Public Function OperaionMode() As enOperationMode


    If UCase$(gAIMS_SERVER_TYPE) = "SERVER" Then
        OperaionMode = enServer
    Else
        OperaionMode = enTerminal
    End If

'    If UCase$(Environ$("AIMS_SERVER_TYPE")) = "SERVER" Then
'        OperaionMode = enServer
'    Else
'        OperaionMode = enTerminal
'    End If
End Function

Public Function ReportGenAt() As String
    
    If OperaionMode = enServer Then
        ReportGenAt = "Gerarated At : Server"
    Else
        ReportGenAt = "Generate At : Terminal [" & gTerminalId & "]"
    End If
    
End Function

Public Sub Dither(frm As Form)
   Dim intLoop As Integer                       ' Counter
       
   ' Set the pen parameters
   frm.DrawStyle = vbInsideSolid
   frm.DrawMode = vbCopyPen
   frm.ScaleMode = vbPixels
   frm.DrawWidth = 8
   frm.ScaleWidth = 256
   
   For intLoop = 0 To 255
      frm.Line (intLoop, 0)-(intLoop - 1, Screen.Height), RGB(0, intLoop, intLoop), B
   Next intLoop
End Sub

Public Sub SetReportFilters(s_Type As String, n_Val As Integer, s_Val As String)

    SQL = SQL & " Insert into tmpReportFilters " & vbCrLf
    SQL = SQL & " (Type,nVal,sVal) " & vbCrLf
    SQL = SQL & " Values ( " & vbCrLf
    SQL = SQL & AQ(s_Type) & vbCrLf
    SQL = SQL & "," & n_Val & vbCrLf
    SQL = SQL & "," & AQ(s_Val) & vbCrLf
    SQL = SQL & ")"
    
    gCnnMst.Execute SQL
    
End Sub

Public Sub ResetReportFilters()

    SQL = " Truncate Table tmpReportFilters "
    
    gCnnMst.Execute SQL
    
End Sub

Public Function GenReportSP(s_SPName As String, s_spPrm() As String)

    Dim tmpStr As String
    Dim iCnt As Integer
    
    tmpStr = " EXEC " + s_SPName
    
    For iCnt = 0 To UBound(s_spPrm)
        tmpStr = tmpStr + AQ(s_spPrm(iCnt)) + ","
    Next
    
    tmpStr = Mid$(tmpStr, 1, Len(tmpStr) - 1)
    
    GenReportSP = tmpStr

End Function


