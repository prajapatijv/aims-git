VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdiMainMenu 
   BackColor       =   &H00F8D9BC&
   ClientHeight    =   5265
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8745
   Icon            =   "mainmenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":116E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":18DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":2724
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":2B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":2E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":32E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":3734
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":3B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":3FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":42F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":4744
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":4B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":4EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":5302
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":5754
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "Sep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Add"
            Object.ToolTipText     =   "Add New Entry"
            Object.Tag             =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Object.ToolTipText     =   "Edit Current Entry"
            Object.Tag             =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Del"
            Object.ToolTipText     =   "Delete Current Entry"
            Object.Tag             =   "Del"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "sep2"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Object.ToolTipText     =   "Save Entry"
            Object.Tag             =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Save"
            Object.ToolTipText     =   "Save And Add New"
            Object.Tag             =   "SaveNAdd"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Object.ToolTipText     =   "Cancel Entry"
            Object.Tag             =   "Cancel"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "sep3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print"
            Object.ToolTipText     =   "Print Entry"
            Object.Tag             =   "Print"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&View"
            Object.ToolTipText     =   "View Entry"
            Object.Tag             =   "View"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "sep4"
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "First"
            Object.ToolTipText     =   "Move First"
            Object.Tag             =   "First"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Object.ToolTipText     =   "Move Next"
            Object.Tag             =   "Next"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prev"
            Object.ToolTipText     =   "Move Previous"
            Object.Tag             =   "Prev"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Last"
            Object.ToolTipText     =   "Move Last"
            Object.Tag             =   "Last"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "sep7"
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            Object.Tag             =   "Find"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "sep5"
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Object.ToolTipText     =   "Exit Entry"
            Object.Tag             =   "Exit"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "sep6"
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialogue 
      Left            =   1200
      Top             =   1365
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   4875
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Visible         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Start"
            TextSave        =   "Start"
            Object.Tag             =   "start"
            Object.ToolTipText     =   "Start"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7937
            MinWidth        =   7937
            Object.ToolTipText     =   "Company Information"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "User Information"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/09/2015"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23:48"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Start 
      Caption         =   ":: &S t a r t ::"
      Begin VB.Menu mnuTrans 
         Caption         =   "&Transctions"
         Begin VB.Menu mnuTransArr1 
            Caption         =   "TrnasArray1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuMasters 
         Caption         =   "&Masters"
         Begin VB.Menu mnuMasterArr1 
            Caption         =   "MasterArray1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuReports 
         Caption         =   "&Reports"
         Begin VB.Menu mnuReportsArr1 
            Caption         =   "ReportsArray"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "&Admin"
         Begin VB.Menu mnuAdminArr1 
            Caption         =   "AdminArray"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSecutiry 
         Caption         =   "&Secutiry"
         Visible         =   0   'False
         Begin VB.Menu mnuSecutiryArr1 
            Caption         =   "SecurityArray"
         End
      End
      Begin VB.Menu mnuUtility 
         Caption         =   "&Utility"
         Begin VB.Menu mnuUtilityArr1 
            Caption         =   "UtilityArray"
            Index           =   0
         End
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSwitchUser 
         Caption         =   "&LogOff"
         Visible         =   0   'False
         Begin VB.Menu mnuSwitchUserArr1 
            Caption         =   "UserArray"
            Index           =   0
         End
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "mdiMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mDbname As String
Dim mFormLoaded As Boolean

Private Sub MDIForm_Activate()
    
    If Not mFormLoaded Then
        
        frmUsrLogin.Show vbModal
        
        mFormLoaded = True
    End If

End Sub

Private Sub MDIForm_DblClick()
    'frmSalesRegister.Show
    'frmInvtrn2.Show
End Sub

Private Sub MDIForm_Load()
On Error GoTo errhndl
MP vbHourglass
    
    VisibleNoVisibleBtn False
    StatBar.Visible = False
    
    LoadTransMenu
    LoadMasterMenu
    LoadReportsMenu
    'LoadSwitchUser
    LoadAdminMenu
    LoadUtilityMenu
    
    
    If Dir(gBgImagePath) <> "" Then
        Me.Picture = LoadPicture(gBgImagePath)
    End If
    
    mFormLoaded = False
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub SetScreenResolution()
On Error GoTo errhndl
MP vbHourglass
    
    SetFormByRes Me

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhndl

    If Len(gUser) > 0 Then
        If Button = vbRightButton Then
            PopupMenu Start
        End If
    End If
    
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MsgBox X, Y
    If Y >= 10000 Then
        StatBar.Visible = True
    Else
        StatBar.Visible = False
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Want to Quit ...??? ", vbYesNo, gUserName) = vbNo Then
        Cancel = 1
    Else
        
'        If IsNeaturalUserMode = False Then
'            BackUpDb gMdbMst, False         'BackUp databae before close
'        End If
        
        End
    End If
End Sub

Private Sub MDIForm_Resize()
    
    On Error Resume Next
    
    If OperaionMode = enTerminal Then
        Me.Caption = " AIMS -> [" & gTerminalId & "]"
    Else
        Me.Caption = " AIMS -> SERVER"
    End If
    
    If IsNeaturalUserMode Then
        Me.Caption = Me.Caption & " |User Neatural Mode|"
    End If
    
End Sub

Private Sub mnuAdminArr1_Click(Index As Integer)
On Error GoTo errhndl
    
    Select Case LCase(mnuAdminArr1(Index).Tag)
        Case "usermast"
            frmUserMast.Show
            
        Case "exportdta"
            With frmServerExport
                .Tag = "export"
                .Show
            End With
            
        Case "importdta"
            With frmServerExport
                .Tag = "import"
                .Show
            End With
            
    End Select
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next

End Sub

Public Sub mnuMasterArr1_Click(Index As Integer)
On Error GoTo errhndl
    
    Select Case LCase(mnuMasterArr1(Index).Tag)
        Case "itmmst"
            With frmItemMast
                .Tag = "itmmst"
                .Show
            End With
        
        Case "itmcatmst"
            With frmGenMst
                .Tag = "itmcatmst"
                .Show
            End With
        
        Case "locationmst"
            With frmGenMst
                .Tag = "locationmst"
                .Show
            End With
        
        Case "sizemst"
            With frmGenMst
                .Tag = "sizemst"
                .Show
            End With

        Case "unitmst"
            With frmGenMst
                .Tag = "unitmst"
                .Show
            End With

        Case "keybrditem"
            frmKeybrdSetup.Show
            
        Case "terconfig"
            frmTerminalMast.Show
        
        Case "eventdef"
            frmEventMast.Show
            
    End Select

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub mnuQuit_Click()
    MDIForm_QueryUnload 0, 0
End Sub

Private Sub mnuReportsArr1_Click(Index As Integer)

Dim SpPrm() As String
Dim formulas() As String

    Select Case LCase(mnuReportsArr1(Index).Tag)
        Case "salreg"
            frmSalesRegister.Show
        Case "rep_salsum"
            With frmRepSales
                .Tag = "rep_salsum"
                .Show
            End With
        
        Case "rep_saldet"
            With frmRepSales
                .Tag = "rep_saldet"
                .Show
            End With
        
        Case "rep_itm"
            With frmRepSales
                .Tag = "rep_itm"
                .Show
            End With
            
        Case "rep_keybrdconfig"
            ReDim SpPrm(1) As String
            ReDim formulas(1) As String
            
            SpPrm(0) = 0        'Keybrd Code
            SpPrm(1) = 0        'Preview Enabled : Always Off : Used for Debug

            formulas(0) = "ReportTitle='Keyboard Configuration Report'"
            formulas(1) = "ReportFilter=''"

            SQL = GenReportSP("rptKeybrdConfiguration", SpPrm)
            
            gCnnMst.Execute SQL

            
            With frmCrviewer
                .ViewReport "keybrdconfig.rpt", SpPrm(), formulas(), 0
                .Tag = "rep_keybrdconfig"
                .Show
            End With

            
        Case LCase("rep_ItmLst")
            With frmRepItem
                .Tag = "rep_ItemList"
                .Show
            End With
            
        Case LCase("rep_ItmLst_MinMaxOrderQty")
            With frmRepItem
                .Tag = "rep_ItmLst_MinMaxOrderQty"
                .Show
            End With
            
        Case LCase("rep_opcl")
            With frmRepStk
                .Tag = "rep_opcl"
                .Show
            End With
            
        Case LCase("rep_opcl2")
            With frmRepStk
                .Tag = "rep_opcl2"
                .Show
            End With
            
        Case LCase("rep_Invsum")
            With frmRepInv
                .Tag = "rep_Invsum"
                .Show
            End With
        
        Case LCase("rep_ItmBarcode")
            With frmRepBarcode
                .Tag = "rep_ItmBarcode"
                .Show
            End With
        
'        Case LCase("rep_Invdet")
'            With frmRepInv
'                .Tag = "rep_Invdet"
'                .Show
'            End With
        
        
        Case Else
            MsgBox "Report Not Found..!!", vbExclamation
            
    End Select


    
End Sub

Private Sub mnuSwitchUserArr1_Click(Index As Integer)
    Select Case LCase(mnuSwitchUserArr1(Index).Tag)
        Case "usrlogin"
            frmUsrLogin.Show vbModal
    End Select
End Sub

Private Sub mnuTransArr1_Click(Index As Integer)
On Error GoTo errhndl
MP vbHourglass
    
    Select Case LCase(mnuTransArr1(Index).Tag)
        Case "saltrn"
            frmPosGui.Show vbModal
        
        Case "invtrn"
            frmInvtrn2.Show
            
    End Select

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub mnuUtilityArr1_Click(Index As Integer)
    Select Case LCase(mnuUtilityArr1(Index).Tag)
        Case "backdb"
            BackUpDb gMdbMst, True
    End Select
End Sub

Private Sub StatBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
On Error GoTo errhndl
'MP vbHourglass
    
    Select Case StatBar.Panels(1).Tag
        Case "start"
            PopupMenu Start, , 0, StatBar.Top
    End Select

'MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errhndl
MP vbHourglass
    
    Dim i As Integer
    
    For i = 1 To TbrMain.Buttons.Count
        TbrMain.Buttons(i).Enabled = False
    Next
    TbrMain.Buttons(btnsave).Enabled = False
    TbrMain.Buttons(btnSaveNAdd).Enabled = False
    TbrMain.Buttons(btnCancel).Enabled = False
    
    Select Case LCase(Button.Tag)
        Case "add"
            If gAdd Then
                TbrMain.Buttons(btnsave).Enabled = True
                TbrMain.Buttons(btnSaveNAdd).Enabled = True
                TbrMain.Buttons(btnCancel).Enabled = True
                
                Screen.ActiveForm.EntryAdd
            End If
        Case "edit"
            If gEdit Then
                TbrMain.Buttons(btnsave).Enabled = True
                TbrMain.Buttons(btnSaveNAdd).Enabled = True
                TbrMain.Buttons(btnCancel).Enabled = True
            
                Screen.ActiveForm.EntryEdit ViewMode.EntryReadWrite
            End If
        Case "del"
            If gDel Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
                
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
            
                Screen.ActiveForm.EntryDelete
            End If
        Case "save"
            For i = 1 To TbrMain.Buttons.Count
                TbrMain.Buttons(i).Enabled = True
            Next
        
            TbrMain.Buttons(btnsave).Enabled = False
            TbrMain.Buttons(btnSaveNAdd).Enabled = False
            TbrMain.Buttons(btnCancel).Enabled = False
        
            Screen.ActiveForm.EntrySave
            
        Case "saveaddnew" 'save and add
            TbrMain.Buttons(btnsave).Enabled = True
            TbrMain.Buttons(btnSaveNAdd).Enabled = True
            TbrMain.Buttons(btnCancel).Enabled = True
            
            'Screen.ActiveForm.EntrySaveNAdd
            
        Case "cancel"
            For i = 1 To TbrMain.Buttons.Count
                TbrMain.Buttons(i).Enabled = True
            Next
        
            TbrMain.Buttons(btnsave).Enabled = False
            TbrMain.Buttons(btnSaveNAdd).Enabled = False
            TbrMain.Buttons(btnCancel).Enabled = False
        
            Screen.ActiveForm.EntryCancel
            
        Case "print"
            If gPrint Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
                
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
            
                Screen.ActiveForm.EntryPrint
            End If
            
        Case "view"
            'If gEdit Then
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                
                TbrMain.Buttons(btnCancel).Enabled = True
                TbrMain.Buttons(btnadd).Enabled = True
                TbrMain.Buttons(btnedit).Enabled = True
                
                Screen.ActiveForm.EntryEdit ViewMode.EntryReadOnly
            'End If
            
        Case "first"
            If gNevigate Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
            
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
            
                Screen.ActiveForm.EntryFirst
            End If
        Case "next"
            If gNevigate Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
            
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
            
                Screen.ActiveForm.EntryNext
            End If
        Case "prev"
            If gNevigate Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
            
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
            
                Screen.ActiveForm.EntryPrev
            End If
        Case "last"
            If gNevigate Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
            
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
                            
                Screen.ActiveForm.EntryLast
            End If
        Case "find"
            If gNevigate Then
                For i = 1 To TbrMain.Buttons.Count
                    TbrMain.Buttons(i).Enabled = True
                Next
            
                TbrMain.Buttons(btnsave).Enabled = False
                TbrMain.Buttons(btnSaveNAdd).Enabled = False
                TbrMain.Buttons(btnCancel).Enabled = False
                            
                Screen.ActiveForm.EntryFind
            End If
            
        Case "exit"
            Screen.ActiveForm.EntryExit
        
        Case "quit"
            EntryQuit
        
'''        Case "menu"
'''            MDIForm_MouseDown vbRightButton, 0, 0, 0
    End Select
    
    TbrMain.Buttons(btnSaveNAdd).Enabled = False
    TbrMain.Buttons(btnExit).Enabled = True
    
    If Screen.ActiveForm.Controls.Count > 0 Then
        If Not (Screen.ActiveForm.ActiveControl Is Nothing) Then
            Screen.ActiveForm.lblMode.Caption = StrConv(Screen.ActiveForm.mEntryMode, vbProperCase)
        End If
    End If
    
MP vbDefault
Exit Sub
errhndl:
    'ErrMsg
    Resume Next
End Sub

Private Sub EntryQuit()
    
End Sub
Private Sub LoadTransMenu()
On Error GoTo errhndl
MP vbHourglass
        
    mnuTransArr1(0).Caption = "&1. Sales Entry"
    mnuTransArr1(0).Tag = "saltrn"
        
    Load mnuTransArr1(1)
    mnuTransArr1(1).Caption = "&2. Stock Inward/Outward Entry"
    mnuTransArr1(1).Tag = "invtrn"
    
'    Load mnuTransArr1(2)
'    mnuTransArr1(2).Caption = "-"
'    mnuTransArr1(2).Tag = "mnudash1"
        
        
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub LoadMasterMenu()
On Error GoTo errhndl
MP vbHourglass
        
    mnuMasterArr1(0).Caption = "&1. Item Master"
    mnuMasterArr1(0).Tag = "itmmst"
    
    Load mnuMasterArr1(1)
    mnuMasterArr1(1).Caption = "&2. Item Categories"
    mnuMasterArr1(1).Tag = "itmcatmst"

    Load mnuMasterArr1(2)
    mnuMasterArr1(2).Caption = "&3. Locations"
    mnuMasterArr1(2).Tag = "locationmst"

    Load mnuMasterArr1(3)
    mnuMasterArr1(3).Caption = "&4. Sizes"
    mnuMasterArr1(3).Tag = "sizemst"

    Load mnuMasterArr1(4)
    mnuMasterArr1(4).Caption = "&5. Units"
    mnuMasterArr1(4).Tag = "unitmst"

    Load mnuMasterArr1(5)
    mnuMasterArr1(5).Caption = "-"
    mnuMasterArr1(5).Tag = "dash1"

    Load mnuMasterArr1(6)
    mnuMasterArr1(6).Caption = "&6. Item Keyboard Setup"
    mnuMasterArr1(6).Tag = "keybrditem"

    Load mnuMasterArr1(7)
    mnuMasterArr1(7).Caption = "&7. Terminal Configuration"
    mnuMasterArr1(7).Tag = "terconfig"

    Load mnuMasterArr1(8)
    mnuMasterArr1(8).Caption = "&8. Events Defination"
    mnuMasterArr1(8).Tag = "eventdef"

    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub LoadReportsMenu()
On Error GoTo errhndl
MP vbHourglass
        
    mnuReportsArr1(0).Caption = "&1. Sales Register"
    mnuReportsArr1(0).Tag = "salreg"
    
    Load mnuReportsArr1(1)
    mnuReportsArr1(1).Caption = "&2. Sales Summary Reports"
    mnuReportsArr1(1).Tag = "rep_salsum"
    
    Load mnuReportsArr1(2)
    mnuReportsArr1(2).Caption = "&3. Sales Detail Report"
    mnuReportsArr1(2).Tag = "rep_saldet"
    
    Load mnuReportsArr1(3)
    mnuReportsArr1(3).Caption = "&4. Itemwise Sales Report"
    mnuReportsArr1(3).Tag = "rep_itm"
    
    Load mnuReportsArr1(4)
    mnuReportsArr1(4).Caption = "-"
    mnuReportsArr1(4).Tag = "mnuRepArrdash2"
        
    Load mnuReportsArr1(5)
    mnuReportsArr1(5).Caption = "&5. Keyboard Configuration Report"
    mnuReportsArr1(5).Tag = "rep_keybrdconfig"
        
    Load mnuReportsArr1(6)
    mnuReportsArr1(6).Caption = "&6. Item List Report"
    mnuReportsArr1(6).Tag = "rep_ItmLst"
        
    Load mnuReportsArr1(7)
    mnuReportsArr1(7).Caption = "&7. Item Min Max Order Quantity Report"
    mnuReportsArr1(7).Tag = "rep_ItmLst_MinMaxOrderQty"
        
    Load mnuReportsArr1(8)
    mnuReportsArr1(8).Caption = "-"
    mnuReportsArr1(8).Tag = "mnuRepArrdash3"
        
    Load mnuReportsArr1(9)
    mnuReportsArr1(9).Caption = "&8. Stock Report"
    mnuReportsArr1(9).Tag = "rep_opcl"
        
    Load mnuReportsArr1(10)
    mnuReportsArr1(10).Caption = "&9. Stock Report - 2"
    mnuReportsArr1(10).Tag = "rep_opcl2"
    
    Load mnuReportsArr1(11)
    mnuReportsArr1(11).Caption = "&A. Inventory Summary Report"
    mnuReportsArr1(11).Tag = "rep_Invsum"
        
    Load mnuReportsArr1(12)
    mnuReportsArr1(12).Caption = "&B. Barcode Label Generator"
    mnuReportsArr1(12).Tag = "rep_ItmBarcode"
        
'    Load mnuReportsArr1(10)
'    mnuReportsArr1(10).Caption = "&A. Inventory Detail Report"
'    mnuReportsArr1(10).Tag = "rep_Invdet"
        
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub LoadSwitchUser()
    
    mnuSwitchUserArr1(0).Caption = "&1. LogOff"
    mnuSwitchUserArr1(0).Tag = "usrlogin"
    
    'Load mnuSwitchUserArr1(1)
End Sub

Private Sub LoadAdminMenu()
On Error GoTo errhndl
MP vbHourglass
        
    mnuAdminArr1(0).Caption = "&1. User Master"
    mnuAdminArr1(0).Tag = "usermast"
        
    Load mnuAdminArr1(1)
    mnuAdminArr1(1).Caption = "&2. Export Data"
    mnuAdminArr1(1).Tag = "exportdta"
        
    Load mnuAdminArr1(2)
    mnuAdminArr1(2).Caption = "&3. Import Data"
    mnuAdminArr1(2).Tag = "importdta"
        
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub LoadUtilityMenu()
On Error GoTo errhndl
MP vbHourglass
        
    mnuUtilityArr1(0).Caption = "&1. BackUp Database"
    mnuUtilityArr1(0).Tag = "BackDB"
            
'    Load mnuUtilityArr1(1)
'    mnuUtilityArr1(1).Caption = "&2. Calc"
'    mnuUtilityArr1(1).Tag = "Calc"
        
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub


