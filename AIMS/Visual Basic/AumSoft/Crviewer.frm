VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCrviewer 
   BackColor       =   &H8000000A&
   Caption         =   "R e p o r t   V i e w e r "
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "Crviewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6735
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   5895
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComctlLib.ImageList ImgListReport 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":0EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":1AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":1DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":20E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":23FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crviewer.frx":2716
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicRepChoi 
      BackColor       =   &H00F8D9BC&
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8715
      ScaleWidth      =   2715
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2775
      Begin MSComctlLib.TreeView tvRepChoi 
         Height          =   10455
         Left            =   0
         TabIndex        =   0
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   18441
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblReport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">>>  Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1200
         TabIndex        =   2
         Top             =   80
         Width           =   1320
      End
   End
   Begin MSComctlLib.Toolbar ReportBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   635
      ButtonWidth     =   2275
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgListReport"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Report Tree"
            Key             =   "Report Tree"
            Object.ToolTipText     =   "Show/Hide Report Tree"
            Object.Tag             =   "Report Tree"
            ImageIndex      =   7
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Group Tree"
            Key             =   "Group Tree"
            Object.ToolTipText     =   "Group Tree"
            Object.Tag             =   "Group Tree"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Setup"
            Key             =   "Print Setup"
            Object.ToolTipText     =   "Display Print Setup Dialogue Box"
            Object.Tag             =   "Print Setup"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Report"
            Object.Tag             =   "Exit"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmCrviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFrmObj As Form
Dim mFormLoaded As Boolean
Dim mReportObj As CRAXDRT.Report

Private Sub CRViewer1_Click()

End Sub

Private Sub Form_Activate()
    
    If Not mFormLoaded Then
    
        Screen.MousePointer = vbHourglass
    
        With PicRepChoi
            .Move 0, ReportBar.Height
            .Height = Val(Me.ScaleHeight) - Val(ReportBar.Height)
            .Visible = False
        End With
    
        With CRViewer1
            .EnableCloseButton = True
            .EnableExportButton = True
            .EnableGroupTree = False
            .EnableRefreshButton = False
            .EnableSearchExpertButton = True
            .EnableSearchControl = True
            .EnableZoomControl = True
            .Zoom 1
            
        End With
            
'        Select Case LCase(Me.Tag)
'            Case LCase("Pattern")
'                AddPatternRepoChoi
'        End Select
        
        Screen.MousePointer = vbDefault

        mFormLoaded = True
    End If
    
End Sub

Private Sub Form_Load()
    If Not (Screen.ActiveForm Is Nothing) Then Set mFrmObj = Screen.ActiveForm
    mFormLoaded = False
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With CRViewer1
        .Top = ReportBar.Height
        .Left = 0                           'IIf(PicRepChoi.Visible, PicRepChoi.Width + 40, 0)
        .Height = Me.ScaleHeight - ReportBar.Height
        .Width = Me.ScaleWidth              '- IIf((PicRepChoi.Visible And ScaleWidth > 0), PicRepChoi.Width, 0)
    End With
    
    CenterFormCaption Me, Me.Caption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub GenerateReport(s_CryRptPath As String, s_spPrm() As String, s_formulas() As String, s_Conum As Integer)

    MP vbHourglass
    
    Dim CryApp As CRAXDDRT.Application
    Dim ReportObj As CRAXDDRT.Report
    Dim DBTables As CRAXDRT.DatabaseTables
    Dim DBTable As CRAXDRT.DatabaseTable
    Dim iCnt As Integer
    Dim Subi As Integer
    Dim Substr As String
    
    Set CryApp = New CRAXDDRT.Application
    Set ReportObj = New CRAXDDRT.Report
        
    ''CryApp.LogOnServerEx "pdsodbc.dll", gSrvName, gMdbMst, gSrvUID, gSrvPwd
     
    Set ReportObj = CryApp.OpenReport(gPathReport & s_CryRptPath)
    ReportObj.DiscardSavedData
    ReportObj.VerifyOnEveryPrint = True
    
    Set DBTables = ReportObj.Database.Tables
    
    iCnt = 0

   If DBTables.Count > 0 Then
        For iCnt = 1 To DBTables.Count
            Set DBTable = DBTables(iCnt)
            DBTable.SetLogOnInfo gSrvName, gMdbMst, gSrvUID, gSrvPwd
            Subi = InStr(1, DBTable.Location, ".", vbTextCompare)
            Substr = gMdbMst & Mid(DBTable.Location, Subi, Len(DBTable.Location))
            DBTable.Location = Substr
        Next
    End If
    'ReportObj.VerifyOnEveryPrint = True

    '---Setting of Report Heading Formulas
    'SetReportHeadings s_Conum, s_RepHeader, s_RepFilter, ReportObj
    'ReportObj.FormulaFields.GetItemByName("ReportTitle").Text = "'Test'"
    Call SetReportFormulas(s_formulas, ReportObj)
    
    Set mReportObj = ReportObj

    ReportObj.EnableParameterPrompting = False
    
    With CRViewer1
        .ReportSource = mReportObj
        .ViewReport
        .Zoom 1
    End With
    
    CenterFormCaption Me, s_CryRptPath
    
    Set CryApp = Nothing
    Set ReportObj = Nothing
    
    MP vbDefault
    
End Sub

Private Sub SetReportFormulas(s_formulas() As String, s_ReportObj As CRAXDRT.Report)

    Dim iCnt As Integer
    Dim arr
    
    If UBound(s_formulas) = 0 Then
        Exit Sub
    End If
        
    For iCnt = 0 To UBound(s_formulas)
        arr = Split(s_formulas(iCnt), "=")
        
        With s_ReportObj.FormulaFields
            .GetItemByName(arr(0)).Text = arr(1)
        End With
    Next
    
End Sub

Private Sub SetSubReports(s_PatentReportObj As CRAXDRT.Report)
    
    Dim xdb As CRAXDRT.Database
    Dim xDBTables As CRAXDRT.DatabaseTables
    Dim xDBTable As CRAXDRT.DatabaseTable
    Dim xReportObj As CRAXDRT.SubreportObject
    Dim xSection As Section
    Dim rptObject As Object
            
    Dim Substr As String
    Dim Subi As Integer
    Dim Subcnt As Integer

    For Each xSection In s_PatentReportObj.Sections
    
        For Each rptObject In xSection.ReportObjects
            If (rptObject.Kind = crSubreportObject) Then
                Set xReportObj = rptObject
                
                xReportObj.OpenSubreport.Database.LogOnServer "pdssql.dll", gSrvName, gMdbMst, gSrvUID, gSrvPwd
                
                Set xdb = xReportObj.OpenSubreport.Database
                Set xDBTables = xdb.Tables
                
                Subi = 0
            
                If xDBTables.Count > 0 Then
                    For Subcnt = 1 To xDBTables.Count
                        Set xDBTable = xDBTables(Subcnt)
                        xDBTable.SetLogOnInfo gSrvName, gMdbMst, gSrvUID, gSrvPwd
                        Subi = InStr(1, xDBTable.Location, ".", vbTextCompare)
                        Substr = gMdbMst & Mid(xDBTable.Location, Subi, Len(xDBTable.Location))
                        xDBTable.Location = Substr
                    Next
                End If
                Set xReportObj = Nothing
            End If
        Next
    
    Next

    Set xdb = Nothing
    Set xSection = Nothing
    Set xReportObj = Nothing
    Set xDBTables = Nothing
    Set xDBTable = Nothing
    Set rptObject = Nothing
        
End Sub

Private Sub SetReportHeadings(s_Conum As Integer, s_Title As String, s_RepFilter As String, s_Repo As CRAXDRT.Report)
    Dim rstMast As ADODB.Recordset
    Dim i As Integer
    SQL = "Select code,[name],Fyear,add1,add2,add3,city,phone,phone1,phone2,mobile,email,www,CstNo,GstNo,PanNo"
    SQL = SQL & " From " & GetDbTable("Compmast", gMdbMst)
    SQL = SQL & " Where code = " & Val(s_Conum)
    
    OpenAdoRst rstMast, SQL
    With rstMast
        If .RecordCount > 0 Then
            For i = 1 To s_Repo.FormulaFields.Count
                Select Case LCase(s_Repo.FormulaFields.Item(i).FormulaFieldName)
                '---Standard Formulas
                Case "company"
                    s_Repo.FormulaFields.GetItemByName("Company").Text = AQ(IfNullThen(.Fields("Name").Value, "")) '& " - " & IfNullThen(.Fields("Fyear").Value, ""))
                Case "address"
                    s_Repo.FormulaFields.GetItemByName("Address").Text = AQ(IfNullThen(.Fields("Add1").Value, "") & "," & IfNullThen(.Fields("Add2").Value, "") & "," & IfNullThen(.Fields("City").Value, ""))
                Case "reporttitle"
                    s_Repo.FormulaFields.GetItemByName("ReportTitle").Text = AQ(s_Title)
                Case "reportfilter"
                    s_Repo.FormulaFields.GetItemByName("ReportFilter").Text = AQ(s_RepFilter)
                    
                '---Custom Formulas
                Case "companyname"
                    s_Repo.FormulaFields.GetItemByName("companyname").Text = AQ(IfNullThen(.Fields("name").Value, ""))
                Case "companyadd1"
                    s_Repo.FormulaFields.GetItemByName("companyadd1").Text = AQ(IfNullThen(.Fields("Add1").Value, ""))
                Case "companyadd2"
                    s_Repo.FormulaFields.GetItemByName("companyadd2").Text = AQ(IfNullThen(.Fields("Add2").Value, ""))
                Case "companyadd3"
                    s_Repo.FormulaFields.GetItemByName("companyadd3").Text = AQ(IfNullThen(.Fields("Add3").Value, ""))
                Case "companycity"
                    s_Repo.FormulaFields.GetItemByName("companycity").Text = AQ(IfNullThen(.Fields("city").Value, ""))
                Case "companyphone"
                    s_Repo.FormulaFields.GetItemByName("companyPhone").Text = AQ(IfNullThen(.Fields("Phone").Value, ""))
                Case "companyphone1"
                    s_Repo.FormulaFields.GetItemByName("companyphone1").Text = AQ(IfNullThen(.Fields("phone1").Value, ""))
                Case "companyphone2"
                    s_Repo.FormulaFields.GetItemByName("companyphone2").Text = AQ(IfNullThen(.Fields("phone2").Value, ""))
                Case "companymobile"
                    s_Repo.FormulaFields.GetItemByName("companymobile").Text = AQ(IfNullThen(.Fields("mobile").Value, ""))
                Case "companyemail"
                    s_Repo.FormulaFields.GetItemByName("companyemail").Text = AQ(IfNullThen(.Fields("email").Value, ""))
                Case "companywww"
                    s_Repo.FormulaFields.GetItemByName("companywww").Text = AQ(IfNullThen(.Fields("www").Value, ""))
                Case "companycstno"
                    s_Repo.FormulaFields.GetItemByName("companycstno").Text = AQ(IfNullThen(.Fields("cstno").Value, ""))
                Case "companygstno"
                    s_Repo.FormulaFields.GetItemByName("companygstno").Text = AQ(IfNullThen(.Fields("gstno").Value, ""))
                Case "companypanno"
                    s_Repo.FormulaFields.GetItemByName("companypanno").Text = AQ(IfNullThen(.Fields("panno").Value, ""))

                End Select
            Next
        End If
    End With
    CloseAdoRst rstMast
    
End Sub

Private Sub HideRepSelection()
    With PicRepChoi
        .Top = ReportBar.Height
        .Visible = False
    End With
    
    Form_Resize
End Sub

Private Sub ShowRepSelection()
    With PicRepChoi
        .Top = ReportBar.Height
        .Visible = True
    End With
    Form_Resize
End Sub

Public Sub ViewReport(s_CryRptPath As String, s_spPrm() As String, s_formulas() As String, s_Conum As Integer)
    GenerateReport s_CryRptPath, s_spPrm, s_formulas, s_Conum
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mReportObj = Nothing
End Sub


Private Sub ReportBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case LCase(Button.Tag)
        Case "report tree"
            If Button.Value = tbrPressed Then
                ShowRepSelection
            Else
                HideRepSelection
            End If
            
        Case "print setup"
            mReportObj.PrinterSetup Me.hWnd
            
            With CRViewer1
                .ReportSource = mReportObj
                .ViewReport
                .Zoom 1
            End With
            
        Case "group tree"
            CRViewer1.EnableGroupTree = (Button.Value = tbrPressed)
        Case "exit"
            Unload Me
            
    End Select
    
End Sub

Private Sub tvRepChoi_Click()
    
    With mFrmObj
        Select Case LCase(tvRepChoi.Nodes(tvRepChoi.SelectedItem.Index).Key)
    
        '---Production And Consumption Reports
            '---Production
            Case "proddate"
                .Tag = "cryProductionDat"
                .cmdPrint_Click
                
        End Select
    
    SetFocusTo tvRepChoi
        
    End With
End Sub

Private Sub tvRepChoi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tvRepChoi_Click
    End Select
End Sub

Private Sub SetSPParams(s_spPrm() As String, ByRef s_Repo As CRAXDRT.Report)
    
    Dim iPrmCnt As Integer
    Dim i As Integer

    On Error Resume Next
    
    iPrmCnt = UBound(s_spPrm) - LBound(s_spPrm)
    
    If iPrmCnt >= 1 Then
    
        For i = 0 To UBound(s_spPrm)
        
            With s_Repo.ParameterFields(i)
                
                Select Case .ValueType
                    Case crNumberField:
                        If Left(s_spPrm(i), 1) = vbNullChar Then
                            .AddCurrentValue 0
                        Else
                            .AddCurrentValue CDbl(s_spPrm(i))
                        End If
                    Case crStringField, crDateField, crDateTimeField:
                            .AddCurrentValue s_spPrm(i)

                    Case crCurrencyField:
                            .AddCurrentValue CCur(s_spPrm(i))
                    Case crBooleanField:
                            .AddCurrentValue CBool(s_spPrm(i))
                End Select
                
            End With
            
        Next
        
    End If

End Sub

Public Function SetFormula(FormulaName, FormulaValue) As Boolean
    On Error Resume Next
    Dim rpt
    
    Set rpt = mReportObj
    rpt.FormulaFields.GetItemByName(Trim(FormulaName)).Text = FormulaValue

End Function

'Public Function ExportFileType(sfileName As String)
'On Error GoTo CheckError
'    CRXReport.EnableParameterPrompting = False
'    CRXReport.DisplayProgressDialog = False
'    CRXReport.MorePrintEngineErrorMessages = mShowMsg
'    CRXReport.ExportOptions.FormatType = crEFTExactRichText
'    CRXReport.ExportOptions.DestinationType = crEDTDiskFile
'    CRXReport.ExportOptions.DiskFileName = sfileName
'    CRXReport.Export False
'
'Exit Function
'CheckError:
'       Err.Raise Err.Number, "cCrystal::Export", Err.Description
'End Function
'
'Public Function ExportToPrinter(Optional iNumberOfCopies As Integer = 1)
'
'On Error GoTo CheckError
'    CRXReport.EnableParameterPrompting = False
'    CRXReport.MorePrintEngineErrorMessages = False
'    CRXReport.MorePrintEngineErrorMessages = True
'    CRXReport.DisplayProgressDialog = False
'
'    If iNumberOfCopies < 1 Then Exit Function
'    CRXReport.PrintOut False, iNumberOfCopies
'
'Exit Function
'CheckError:
'       If Err.Description <> ERROR_DETECTED_BY_DATABASE_DLL Then
'             Err.Raise Err.Number, "cCrystal::ExportToPrinter", Err.Description
'       End If
'
'End Function
'
'Public Function SetDataSource(DSN As Object)
'    CRXReport.Database.SetDataSource DSN
'    CRXReport.Database.Verify
'End Function

Private Sub AddPatternRepoChoi()
    Dim NodeP As Node
    Dim NodeC As Node
    Dim i As Integer
    
    With tvRepChoi
        .Style = tvwTreelinesPlusMinusText
        .Indentation = 150
        .Font.Size = 10
        .FullRowSelect = True
        .LineStyle = tvwRootLines
        .HotTracking = True
        .Style = tvwTreelinesPlusMinusPictureText
        
        Set NodeP = .Nodes.Add(, , "Pattern", "Pattern")
        NodeP.Expanded = True
        
        
        Set NodeC = .Nodes.Add("Pattern", tvwChild, "PatPallete", "PlateWise")
        Set NodeC = .Nodes.Add("Pattern", tvwChild, "PatParty", "PartyWise")
        Set NodeC = .Nodes.Add("Pattern", tvwChild, "PatLoc", "LocationWise")
        Set NodeC = .Nodes.Add("Pattern", tvwChild, "PatLed", "Pattern Ledger")
        Set NodeC = .Nodes.Add("Pattern", tvwChild, "PatPartyStat", "PartyWise Statement")
        
        For i = 1 To .Nodes.Count
            If .Nodes(i).Children > 0 Then
                .Nodes(i).ForeColor = vbBlack
                .Nodes(i).BackColor = &H8000000A  '&HAFD7DC
                .Nodes(i).Bold = True
            Else
                .Nodes(i).ForeColor = vbBlue
            End If
        Next
        
    End With
End Sub


