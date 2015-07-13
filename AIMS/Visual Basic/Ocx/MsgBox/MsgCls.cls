VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgCls"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum MsgMode
    jvOkOnly = 0
    jvOkCancel = 1
    jvOkCancelRetry = 2
    jvOkCancelAbort = 3
    jvYesNo = 4
    jvYesNoCancel = 5
    jvYesNoRetry = 6
    jvAbortRetryCancel = 7
End Enum

Public Enum ReplyMode
    jvOk = 0
    jvCancel = 1
    jvRetry = 2
    jvAbort = 3
    jvYes = 4
    jvNo = 5
End Enum

Public Enum DefButton
    Btn1 = 0
    Btn2 = 1
    Btn3 = 2
End Enum

Public Enum AllignConst
    jvLeft = 0
    jvRight = 1
    jvCenter = 2
End Enum

Public Function MsgBx(s_Msg As String, Optional s_MsgMode As MsgMode, Optional s_defaBtn As DefButton = 1, Optional s_FontSize As Integer = 10, Optional s_Allign As AllignConst = jvLeft, Optional IsBlink As Boolean = False, Optional s_Interval As Integer = 300, Optional s_Header As String = "") As ReplyMode
    Dialog.Button1.Visible = False
    Dialog.Button2.Visible = False
    Dialog.Button3.Visible = False
        
        
    Select Case s_MsgMode
        Case jvOkOnly
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = False
            Dialog.Button3.Visible = False
            Dialog.Button1.Caption = "OK"
        
        Case jvOkCancel
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = False
            Dialog.Button1.Caption = "OK"
            Dialog.Button2.Caption = "Cancel"
        
        Case jvOkCancelRetry
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = True
            Dialog.Button1.Caption = "OK"
            Dialog.Button2.Caption = "Cancel"
            Dialog.Button3.Caption = "Retry"
            
        Case jvOkCancelAbort
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = True
            Dialog.Button1.Caption = "OK"
            Dialog.Button2.Caption = "Cancel"
            Dialog.Button3.Caption = "Abort"
        
        Case jvYesNo
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = False
            Dialog.Button1.Caption = "Yes"
            Dialog.Button2.Caption = "No"

        Case jvYesNoCancel
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = True
            Dialog.Button1.Caption = "Yes"
            Dialog.Button2.Caption = "No"
            Dialog.Button3.Caption = "cancel"
            
        Case jvYesNoRetry
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = True
            Dialog.Button1.Caption = "Yes"
            Dialog.Button2.Caption = "No"
            Dialog.Button3.Caption = "Retry"
        
        Case jvAbortRetryCancel
            Dialog.Button1.Visible = True
            Dialog.Button2.Visible = True
            Dialog.Button3.Visible = True
            Dialog.Button1.Caption = "Abort"
            Dialog.Button2.Caption = "Retry"
            Dialog.Button3.Caption = "Cancel"

    End Select
    
    
    
'''set defult values
    gDefaBtn = s_defaBtn
    
    If IsBlink Then Dialog.Blink.Interval = s_Interval
        
        
    
    Dialog.lblMsg.FontSize = s_FontSize
    Dialog.lblMsg.Alignment = s_Allign
    Dialog.lblHeadB.Caption = s_Header
    Dialog.lblHeadF.Caption = s_Header
    Dialog.lblMsg.Caption = s_Msg
    gMsgStr = s_Msg
    
    Beep
    '''
    Dialog.Show vbModal
    
    Select Case LCase(gBtnClicked)
        Case LCase("jvOk")
            MsgBx = 0
        Case LCase("jvCancel")
            MsgBx = 1
        Case LCase("jvRetry")
            MsgBx = 2
        Case LCase("jvAbort")
            MsgBx = 3
        Case LCase("jvYes")
            MsgBx = 4
        Case LCase("jvNo")
            MsgBx = 5
    End Select
    
    Beep
    
End Function

