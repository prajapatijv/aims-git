VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmKeybrdSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Setup"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "KeybrdSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11355
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   13361
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Master >>"
      TabPicture(0)   =   "KeybrdSetup.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeMaster"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fmeMaster 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   11115
         Begin MSComctlLib.TreeView tvCategories 
            Height          =   5535
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   9763
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin VB.CheckBox chkActv_Fg 
            Height          =   255
            Left            =   2760
            TabIndex        =   1
            Top             =   420
            Width           =   1215
         End
         Begin VB.Frame fmeValues1 
            Height          =   5655
            Left            =   2310
            TabIndex        =   9
            Top             =   1320
            Width           =   8655
            Begin VB.CommandButton cmdSwap 
               Caption         =   "Down"
               Height          =   735
               Index           =   1
               Left            =   7200
               TabIndex        =   16
               Top             =   2400
               Width           =   1335
            End
            Begin VB.CommandButton cmdSwap 
               Caption         =   "Up"
               Height          =   735
               Index           =   0
               Left            =   7200
               TabIndex        =   15
               Top             =   1680
               Width           =   1335
            End
            Begin VB.CommandButton UnselectAll 
               Caption         =   "Unselect All"
               Height          =   495
               Left            =   7200
               TabIndex        =   12
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton cmdSelectAll 
               Caption         =   "Select All"
               Height          =   495
               Left            =   7200
               TabIndex        =   11
               Top             =   360
               Width           =   1335
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
               Height          =   4815
               Left            =   120
               TabIndex        =   10
               Top             =   720
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   8493
               _Version        =   393216
               ForeColor       =   8388608
               Rows            =   1
               FixedRows       =   0
               ForeColorFixed  =   8388608
               BackColorSel    =   16308668
               ForeColorSel    =   8388608
               BackColorBkg    =   14482428
               FocusRect       =   0
               SelectionMode   =   1
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHeader 
               Height          =   495
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   873
               _Version        =   393216
               ForeColor       =   8388608
               ForeColorFixed  =   8388608
               BackColorSel    =   16308668
               ForeColorSel    =   8388608
               BackColorBkg    =   14482428
               FocusRect       =   0
               SelectionMode   =   1
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin CommCtrls.ItxtBox txtCode 
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            Enabled         =   0   'False
            BackColor       =   -2147483643
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
         Begin CommCtrls.CtxtBox txtName 
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   840
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
            Alignment       =   0
            MaxLength       =   60
            AutoCaps        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AllowNull       =   -1  'True
         End
         Begin VB.Label lblMode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mode"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   10320
            TabIndex        =   7
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name*"
            Height          =   240
            Left            =   360
            TabIndex        =   5
            Top             =   900
            Width           =   630
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code*"
            Height          =   240
            Left            =   360
            TabIndex        =   4
            Top             =   420
            Width           =   570
         End
      End
   End
   Begin HlpN.HlpNCode hlpFind 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmKeybrdSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mEntryMode As String
Public mActCtrl As Control

Dim m_CategoryId As Integer

Const dColCheck = 0
Const dColName = 1
Const dColCode = 2
Const dColCheckDb = 3

Public Sub EntryAdd()
On Error GoTo errhndl
MP vbHourglass
   
    mEntryMode = "add"
    ClearScreen
    EnableDisable True
    txtCode.Text = GetMaxCode("KeybrdSetup", True, "Code", gCnnMst)
    chkActv_Fg.Value = 1
    SetFocusTo txtName
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryEdit(iViewMode As ViewMode)
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtCode.Text) <= 0 Then
        MsgBox "No Record Selected For Edit ", vbCritical
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnCancel)
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    If iViewMode = EntryReadWrite Then
        mEntryMode = "edit"
    Else
        mEntryMode = "view"
    End If
    
    EnableDisable True
    SetFocusTo txtName

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryDelete()
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtCode.Text) <= 0 Then
        MsgBox "No Record Selected For Delete ", vbCritical
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnCancel)
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    If MsgBox("Want to Delete EntryNo " & txtCode.Text & "..???", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    mEntryMode = "delete"
    
    gCnnMst.BeginTrans
    
        SQL = "Delete from " & GetDbTable("KeybrdSetup", gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Code = " & Val(txtCode.Text)
        
        SQL = SQL & " Delete from " & GetDbTable("KeybrdItem", gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Keybrd_code = " & Val(txtCode.Text)
    
        gCnnMst.Execute SQL
    
    gCnnMst.CommitTrans
    
    MsgBox "Entry No : " & txtCode.Text & " Deleted ", vbInformation
        
    EntryLast
    
    SetFocusTo SSTab1
    Exit Sub
    
MP vbDefault
Exit Sub
errhndl:
    gCnnMst.RollbackTrans
    
    ErrMsg
    Resume Next
End Sub

Public Sub EntrySave()
On Error GoTo errhndl
MP vbHourglass
    
    If Not ValidateControl Then
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnadd)
        'SetFocusTo txtCode
        Exit Sub
    End If
    
    mEntryMode = "save"
    SaveInTmp
    EnableDisable False
    SetFocusTo SSTab1
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntrySaveNAdd()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "savenadd"
    EntrySave
    EntryAdd

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryCancel()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "cancel"
    ClearScreen
    EnableDisable False

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryPrint()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "print"

    Dim SpPrm(0) As String
    Dim formulas(1) As String
    
    SpPrm(0) = txtCode.Text       'Keybrd Code
    
    formulas(0) = "ReportTitle='Keyboard Configuration Report'"
    If Val(txtCode.Text) > 0 Then
        formulas(1) = "ReportFilter=" & "'" & "KeyBoard : " & txtName.Text & "[" & txtCode.Text & "]" & "'"
    Else
        formulas(1) = "ReportFilter=''"
    End If
    
    formulas(0) = "ReportTitle='Inventory Format Report'"
    
    SQL = GenReportSP("rptInventoryDetail", SpPrm)
    gCnnMst.Execute SQL

    With frmCrviewer
        .ViewReport "keybrdconfig.rpt", SpPrm(), formulas(), 0
        .Tag = "rep_keybrdconfig"
        .Show
    End With


MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryFirst()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "first"
    EnableDisable False
    Nevigate MoveFirst

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryNext()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "next"
    EnableDisable False
    Nevigate MoveNext

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryPrev()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "prev"
    EnableDisable False
    Nevigate MovePrev

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryLast()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "last"
    EnableDisable False
    Nevigate MoveLast

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryFind()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "find"
    EnableDisable False
        
    With hlpFind
        .CodeText = ""
        .NameText = ""
        .Visible = True
        .SetAdoConnStr = gCnnMst
        .TableName = GetDbTable("KeybrdSetup", gMdbMst)
        .FieldList = "Code,name"
        .CodeField = "Code"
        .NameField = "name"
        .SetFocus
        .ShowHelp
    End With
    
    txtCode.Text = Val(hlpFind.CodeText)
    Nevigate MoveTo
    
    hlpFind.Visible = False
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryExit()
    MP vbDefault
    Unload Me
End Sub

Private Sub SetTextBoxes()
        
    hlpFind.Visible = False
    
    VisibleNoVisibleBtn True
    
    SetMsfDetail
    
    AddCategories

    SetActiveModeNControl mEntryMode
    
    CenterFrmChild Me
        
End Sub

Private Sub cmdSelectAll_Click()
    SetFlexFixedColCheckBoxes msfDetail, dColCheck, True, 0
End Sub

Private Sub cmdSwap_Click(Index As Integer)
    SwapItem Index
End Sub

Private Sub tvCategories_Click()
    If tvCategories.SelectedItem.Key <> tvCategories.SelectedItem.Root Then
        m_CategoryId = Replace(tvCategories.SelectedItem.Key, "k", "")
        LoadItemKeybrd m_CategoryId
    End If
End Sub

Private Sub UnselectAll_Click()
    SetFlexFixedColCheckBoxes msfDetail, dColCheck, False, 0
End Sub

Private Sub Form_Activate()
On Error GoTo errhndl
MP vbHourglass
    
    SetFormCaption
    SetTextBoxes

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub SaveInTmp()
On Error GoTo errhndl
MP vbHourglass
        
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from " & GetDbTable("KeybrdSetup", gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Code = " & Val(txtCode.Text)
        gCnnMst.Execute SQL
    End If
    
    SQL = "Insert into KeybrdSetup ("
    SQL = SQL & " code"
    SQL = SQL & ",name"
    SQL = SQL & ",actv_fg"
    
    SQL = SQL & ", dtadat"
    SQL = SQL & ", dtatim"
    SQL = SQL & ", dtausr"
    
    SQL = SQL & ",Trng_fg"

    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtCode.Text)
    SQL = SQL & "," & AQ(txtName.Text)
    SQL = SQL & "," & Val(chkActv_Fg.Value)

    SQL = SQL & "," & ConvDatSql(Date)
    SQL = SQL & "," & AQ(DtaTime)
    SQL = SQL & "," & AQ(gUser)
    
    SQL = SQL & "," & IsTrainingMode
    
    SQL = SQL & ")"
    
    gCnnMst.Execute SQL
    
    SaveItemKbdLink
    
    MsgBox "Entry No : " & txtCode.Text & " Saved ", vbInformation
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg

End Sub

Private Sub SaveItemKbdLink()
On Error GoTo errhndl
MP vbHourglass

    Dim i As Integer
    
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from " & GetDbTable("KeybrdItem", gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And itm_code in "
        SQL = SQL & " ( "
        SQL = SQL & "     select itm_code"
        SQL = SQL & "     from KeybrdItem"
        SQL = SQL & "     Inner Join Items on (Items.Code = KeybrdItem.itm_code)     "
        SQL = SQL & "     where Category_Id = " & m_CategoryId
        SQL = SQL & " ) "
        gCnnMst.Execute SQL
    End If
    
    SQL = ""
    
    With msfDetail
        For i = 0 To .Rows - 1
            If .TextMatrix(i, dColCheck) = "a" Then
                SQL = SQL & "Insert into KeybrdItem ("
                SQL = SQL & " seq"
                SQL = SQL & ",itm_code"
                SQL = SQL & ",actv_fg"
                SQL = SQL & ",keybrd_code"
                
                SQL = SQL & ", dtadat"
                SQL = SQL & ", dtatim"
                SQL = SQL & ", dtausr"
                
                SQL = SQL & ",Trng_fg"
                
                SQL = SQL & " ) Values ("
                
                SQL = SQL & i
                SQL = SQL & "," & Val(.TextMatrix(i, dColCode))
                SQL = SQL & "," & Val(chkActv_Fg.Value)
                SQL = SQL & "," & Val(txtCode.Text)
            
                SQL = SQL & "," & ConvDatSql(Date)
                SQL = SQL & "," & AQ(DtaTime)
                SQL = SQL & "," & AQ(gUser)
                
                SQL = SQL & "," & IsTrainingMode
                
                SQL = SQL & ")" & vbCrLf
            End If
        Next
    End With
    
    If Trim$(SQL) <> "" Then
        gCnnMst.Execute SQL
    End If

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    
End Sub


Public Sub EnableDisable(s_Enable As Boolean)
   fmeMaster.Enabled = s_Enable
End Sub

Private Sub Nevigate(s_Mode As Nevigate)
On Error GoTo errhndl
MP vbHourglass
    
    Dim rsttmp As ADODB.Recordset
    Dim rsttmpCategory As ADODB.Recordset
    
    Select Case s_Mode
        Case MoveFirst
            SQL = "Select top 1 * from " & GetDbTable("KeybrdSetup", gMdbMst) & " Order by Code"
        Case MoveNext
            SQL = "Select top 1 * from " & GetDbTable("KeybrdSetup", gMdbMst) & " Where Code > " & Val(txtCode.Text)
        Case MovePrev
            SQL = "Select top 1 * from " & GetDbTable("KeybrdSetup", gMdbMst) & " where Code < " & Val(txtCode.Text) & " order by Code desc"
        Case MoveLast
            SQL = "Select top 1 * from " & GetDbTable("KeybrdSetup", gMdbMst) & " Order by Code Desc"
        Case MoveTo
            SQL = "Select top 1 * from " & GetDbTable("KeybrdSetup", gMdbMst) & " Where Code=" & Val(txtCode.Text)
    End Select
    
    OpenAdoRst rsttmp, SQL, , , , gCnnMst
    
    With rsttmp
        If .RecordCount > 0 Then
            AdoRsRead rsttmp
            tvCategories.Nodes(2).Selected = True
            tvCategories_Click
        End If
    End With
    
    rsttmp.Close
    Set rsttmp = Nothing

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, Shift, False
End Sub

Private Sub Form_Load()
On Error GoTo errhndl
MP vbHourglass
    
    GrabActiveControl

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_Resize()
    CenterFormCaption Me, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mEntryMode = ""
    Set mActCtrl = Nothing
    VisibleNoVisibleBtn False, True
End Sub

Private Sub SetFormCaption()
On Error GoTo errhndl
MP vbHourglass
    
    SSTab1.Caption = "Keyboard Setup"

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub LoadItemKeybrd(s_CategoryId As Integer)
    Dim i As Integer
    Dim iRow As Integer
    Dim mRow As Integer
    
    Dim rst As ADODB.Recordset
    
    SQL = "Select "
    SQL = SQL & " (Case When Sizes.shortName is null "
    SQL = SQL & "     Then Items.[shortname] "
    SQL = SQL & "     Else Items.[shortname] + '(' + Sizes.shortName + ')'"
    SQL = SQL & " End) as ShortName"
    
    SQL = SQL & ",Items.code        as Code"
    
    SQL = SQL & ",(case when KeybrdItem.itm_code is null then 'r' else 'a' end) as Checkbx"
    SQL = SQL & ",Isnull(KeybrdItem.seq,9999) as seq"
    SQL = SQL & " From " & GetDbTable("Items", gMdbMst)
    
    SQL = SQL & " Left Join " & GetDbTable("KeybrdItem", gMdbMst) & ""
    SQL = SQL & " On (KeybrdItem.itm_code = Items.code"
    SQL = SQL & " And KeybrdItem.keybrd_code = " & Val(txtCode.Text) & ")"
    
    SQL = SQL & " Left Join " & GetDbTable("Sizes", gMdbMst) & ""
    SQL = SQL & " On (Sizes.code = Items.size_id)"
    
    SQL = SQL & " Where Items.actv_fg = 1"
    SQL = SQL & " And Items.category_id = " & Val(s_CategoryId)
    SQL = SQL & " Order by 4"
    
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    Set msfDetail.Recordset = rst
    
    SetMsfDetail
    
    With msfDetail
        .Redraw = False
        mRow = .Row
        If .Rows <> 0 Then
            For iRow = 0 To .Rows - 1
                .Col = dColCheck
                .Row = iRow
                .CellFontName = "Marlett"
                .TextMatrix(iRow, dColCheck) = .TextMatrix(iRow, dColCheckDb)
                
                .Col = dColName
                .CellFontName = gGujaratiFontName
                .CellFontSize = 12
            Next
        
            .Row = 0
            .Col = 1
            .ColSel = .Cols - 2
        End If
        .Redraw = True
    End With
    
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub SetMsfDetail()
On Error GoTo errhndl
MP vbHourglass
    
    'Set Header Grid------------------------------------------------------------------
    With mshHeader
        .RowHeight(0) = 500
        .Cols = 4
        .ScrollBars = flexScrollBarNone
        .Font.Bold = True
        .Enabled = False
        
        .ColAlignment(dColCheck) = flexAlignCenterCenter
        .ColWidth(dColCheck) = 400
        
        .ColWidth(dColCode) = 1260
        .ColAlignment(dColCode) = flexAlignLeftCenter
            
        .ColWidth(dColName) = 5300
        .ColAlignment(dColName) = flexAlignLeftCenter
        
        .ColWidth(dColCheckDb) = 0    ''Checked Status from DB
        
        .TextMatrix(0, dColCheck) = ""
        .TextMatrix(0, dColCode) = "Code"
        .TextMatrix(0, dColName) = "Item Name"
        .TextMatrix(0, dColCheckDb) = "CheckDb"      'Not visible
    End With
    '---------------------------------------------------------------------------------
    
    'Set Detail Grid------------------------------------------------------------------
    With msfDetail
        .FixedCols = 1
        .FixedRows = 0
        .Cols = 4
        .RowHeightMin = 350
        .Font.Bold = True
        .ScrollBars = flexScrollBarVertical
        
        .ColAlignment(dColCheck) = flexAlignCenterCenter
        .ColWidth(dColCheck) = 400
        
        .ColWidth(dColCode) = 1000
        .ColAlignment(dColCode) = flexAlignLeftCenter
            
        .ColWidth(dColName) = 5300
        .ColAlignment(dColName) = flexAlignLeftCenter

        .ColWidth(dColCheckDb) = 0    ''Checked Status from DB
        
        .Move mshHeader.Left, mshHeader.Top + mshHeader.RowHeight(0) - 20
    End With

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub msfDetail_Click()

    Dim iRow As Integer
    Dim mRow As Integer

    With msfDetail
        If .Rows = 0 Then Exit Sub
        mRow = .Row
        For iRow = 0 To .Rows - 1
            If .MouseCol <> 0 Then Exit Sub
            .Col = dColCheck
            .CellFontName = "Marlett"
            If mRow = iRow Then
                If .TextMatrix(iRow, dColCheck) = "a" Then
                    .TextMatrix(iRow, dColCheck) = "r"
                Else
                    .TextMatrix(iRow, dColCheck) = "a"
                End If
            End If
        Next
        .Col = 1
    End With

End Sub

Private Sub AddCategories()
    Dim NodeP As Node
    Dim NodeC As Node
    Dim i As Integer
    
    Dim rst As ADODB.Recordset
    
    With tvCategories
        .Style = tvwTreelinesPlusMinusText
        .Indentation = 150
        .Font.Name = gGujaratiFontName
        .Font.Size = 12
        .FullRowSelect = True
        .LineStyle = tvwRootLines
        .HotTracking = True
        .Style = tvwTreelinesPlusMinusPictureText
        
        '---Categories----------------------------------------------------------------
        Set NodeP = .Nodes.Add(, , "Categories", "fuxudhe")     'Catetory
        NodeP.Expanded = True
        
        SQL = "Select code,[shortname] from Categories where actv_fg = 1"
        OpenAdoRst rst, SQL, , , , gCnnMst
           
        While Not rst.EOF
            Set NodeC = .Nodes.Add("Categories", tvwChild, "k" & rst.Fields("code").Value, rst.Fields("shortname").Value)
            rst.MoveNext
        Wend
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
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub SwapItem(s_SwapMode As Integer)
On Error GoTo errhndl

    Dim tempCheck As String
    Dim tempCode As Integer
    Dim tempName As String
    Dim tempCheckdb As String
    
    Dim mSrcRow As Integer
    Dim mDestRow As Integer
    
    mSrcRow = msfDetail.Row
        
    If msfDetail.TextMatrix(mSrcRow, dColCheck) = "r" Then
        Exit Sub
    End If
    
    If s_SwapMode = 1 Then
        mDestRow = msfDetail.Row + 1
        If msfDetail.Rows = mDestRow Then Exit Sub
    Else
        If msfDetail.Row = 0 Then Exit Sub
        mDestRow = msfDetail.Row - 1
    End If
    
    With msfDetail
        tempCheck = .TextMatrix(mSrcRow, dColCheck)
        tempCode = .TextMatrix(mSrcRow, dColCode)
        tempName = .TextMatrix(mSrcRow, dColName)
        tempCheckdb = .TextMatrix(mSrcRow, dColCheckDb)
            
        .Row = mSrcRow
        .TextMatrix(mSrcRow, dColCheck) = .TextMatrix(mDestRow, dColCheck)
        .TextMatrix(mSrcRow, dColCode) = .TextMatrix(mDestRow, dColCode)
        .TextMatrix(mSrcRow, dColName) = .TextMatrix(mDestRow, dColName)
        .TextMatrix(mSrcRow, dColCheckDb) = .TextMatrix(mDestRow, dColCheckDb)
            
        .Row = mDestRow
        .TextMatrix(mDestRow, dColCheck) = tempCheck
        .TextMatrix(mDestRow, dColCode) = tempCode
        .TextMatrix(mDestRow, dColName) = tempName
        .TextMatrix(mDestRow, dColCheckDb) = tempCheckdb
        
    End With

Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

