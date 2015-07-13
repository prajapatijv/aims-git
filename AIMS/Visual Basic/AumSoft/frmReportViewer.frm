VERSION 5.00
Begin VB.Form frmReportViewer 
   Caption         =   "ReportViewer"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ViewReport(s_RepoName As String, s_spPrm() As String, s_formulas() As String)
    
On Error GoTo errorhandler

    Dim mReportFile As String
    mReportFile = Trim$(gPathReport) & Trim$(s_RepoName)
    
'    With CryRpt
'        If IsFile(mReportFile) Then
'            .ReportFileName = mReportFile
'
'             SetReportParams s_spPrm()
'             SetReportFormulas s_formulas()
'
'            .DiscardSavedData = True
'            .WindowState = crptMaximized
'            .WindowTitle = mReportFile
'            .Action = 1
'        Else
'            MsgBox "Report Not Found " & vbCrLf & vbCrLf & mReportFile, vbCritical
'        End If
'    End With
    
    Exit Sub
    
errorhandler:
    ErrMsg
    Resume Next
End Sub

Private Sub SetReportParams(s_spPrm() As String)

    Dim iCnt As Integer
    
'    If UBound(s_spPrm) = 0 Then
'        Exit Sub
'    End If
        
    For iCnt = 0 To UBound(s_spPrm)
'        With CryRpt
'            .StoredProcParam(iCnt) = s_spPrm(iCnt)
'        End With
    Next
    
End Sub

Private Sub SetReportFormulas(s_formulas() As String)

    Dim iCnt As Integer
    
    If UBound(s_formulas) = 0 Then
        Exit Sub
    End If
        
    For iCnt = 0 To UBound(s_formulas)
'        With CryRpt
'            .formulas(iCnt) = s_formulas(iCnt)
'        End With
    Next
    
End Sub


Private Sub SetReportDefaultParams()

    Select Case Me.Tag
        Case "rep_keybrdconfig"
        
        Case Else
    End Select
    
End Sub

