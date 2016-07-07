VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmInvtrn2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Material Inward/outward Entry"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2
   Icon            =   "Invtrn2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   17806
      _Version        =   393216
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
      TabCaption(0)   =   "Inward/Outward Entry"
      TabPicture(0)   =   "Invtrn2.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTabDetail"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmeTotals"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fmeRecDetail"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FmeCompanyDetail"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   -600
         Width           =   1455
      End
      Begin VB.Frame FmeCompanyDetail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   330
         Width           =   12015
         Begin CommCtrls.ItxtBox txtVno 
            Height          =   375
            Left            =   10800
            TabIndex        =   1
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   240
            TabIndex        =   26
            Top             =   300
            Width           =   600
         End
         Begin VB.Label lblVno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vno*"
            Height          =   240
            Left            =   10200
            TabIndex        =   10
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame fmeRecDetail 
         Height          =   1695
         Left            =   120
         TabIndex        =   20
         Top             =   945
         Width           =   12015
         Begin VB.CommandButton cmdViewDocument 
            Caption         =   "View document"
            Height          =   375
            Left            =   3720
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton cmdFile 
            Caption         =   "Upload document"
            Height          =   375
            Left            =   1680
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1935
         End
         Begin MSComDlg.CommonDialog filedialogue 
            Left            =   5160
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.ComboBox cboItemCategory 
            Height          =   360
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   3255
         End
         Begin VB.ComboBox cmbType 
            Height          =   360
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   3255
         End
         Begin CommCtrls.mskDat mskRec_Dat 
            Height          =   375
            Left            =   10800
            TabIndex        =   6
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            AllowNull       =   -1  'True
         End
         Begin CommCtrls.CtxtBox txtDoc_No 
            Height          =   375
            Left            =   6960
            TabIndex        =   5
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Alignment       =   0
            MaxLength       =   30
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
         Begin HlpN.HlpNCode hlpTerminalcode 
            Height          =   375
            Left            =   6960
            TabIndex        =   4
            Top             =   240
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Name"
            Height          =   240
            Left            =   240
            TabIndex        =   31
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label lblLocation_id 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terminal"
            Height          =   240
            Left            =   6000
            TabIndex        =   28
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Category"
            Height          =   240
            Left            =   240
            TabIndex        =   27
            Top             =   780
            Width           =   1245
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   240
            Left            =   240
            TabIndex        =   25
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lblReqNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DocNo"
            Height          =   240
            Left            =   6000
            TabIndex        =   11
            Top             =   780
            Width           =   645
         End
         Begin VB.Label lblReqDat 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            Height          =   240
            Left            =   10200
            TabIndex        =   12
            Top             =   780
            Width           =   435
         End
      End
      Begin VB.Frame fmeTotals 
         Height          =   1245
         Left            =   135
         TabIndex        =   21
         Top             =   8700
         Width           =   12045
         Begin CommCtrls.CtxtBox txtRemarks 
            Height          =   855
            Left            =   1320
            TabIndex        =   9
            Top             =   195
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   1508
            BackColor       =   14482428
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
         Begin CommCtrls.NTxtBox txtTotRecQty 
            Height          =   375
            Left            =   8640
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   195
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Locked          =   -1  'True
            BackColor       =   14482428
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
         Begin CommCtrls.NTxtBox txtItemTot 
            Height          =   375
            Left            =   8640
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   675
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Locked          =   -1  'True
            BackColor       =   14482428
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
         Begin VB.Label lblRemarks 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblItemTot 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            Height          =   240
            Left            =   7320
            TabIndex        =   16
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label lblTotRecQty 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Qty"
            Height          =   240
            Left            =   7320
            TabIndex        =   14
            Top             =   240
            Width           =   810
         End
      End
      Begin TabDlg.SSTab SSTabDetail 
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   12045
         _ExtentX        =   21246
         _ExtentY        =   10398
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   2
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
         TabCaption(0)   =   "&1. Item Detail"
         TabPicture(0)   =   "Invtrn2.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fmeMsfdetail"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame fmeMsfdetail 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5415
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   11895
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
               Height          =   5115
               Left            =   120
               TabIndex        =   8
               Top             =   210
               Width           =   11655
               _ExtentX        =   20558
               _ExtentY        =   9022
               _Version        =   393216
               Cols            =   1
               FixedCols       =   0
               BackColorSel    =   16308668
               ForeColorSel    =   0
               FocusRect       =   0
               HighLight       =   2
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   1
            End
            Begin VB.Label lblEccNo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4920
               TabIndex        =   24
               Top             =   840
               Width           =   975
            End
         End
      End
   End
   Begin HlpN.HlpNCode hlpFind 
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmInvtrn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mEntryMode As String
Public mActCtrl As Control

Dim i As Integer
Dim zerobarredTypes() As String

Const dColItm_id = 0
Const dColItm_name = 1
Const dColQty = 2
Const dColUnit = 3
Const dColrtl_rpc = 4
Const dColAmt = 5
Dim showTerminalCodeForInwardOutwardTypeList() As String
Dim documentPath As String

Public Sub EntryAdd()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "add"
    ClearScreen
    ClearMsf msfDetail
    SSTabDetail.Tab = 0

    
    EnableDisable True
    SetMsfDetail msfDetail
    txtVno.Text = GetMaxVno("Invtrn")
    
    cboItemCategory.Enabled = True
    cmbType.Enabled = True
    SetFocusTo cmbType
        
    mskRec_Dat.Text = Date
    documentPath = ""
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
    
End Sub

Public Sub EntryEdit(iViewMode As ViewMode)
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtVno.Text) <= 0 Then
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
    
    SetFocusTo txtDoc_No

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryDelete()
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtVno.Text) <= 0 Then
        MsgBox "No Record Selected For Delete ", vbCritical
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnCancel)
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    If MsgBox("Want to Delete EntryNo " & txtVno.Text & "..???", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    mEntryMode = "delete"
    
    SQL = "Delete from Invtrn"
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Vno= " & Val(txtVno.Text)
    gCnnMst.Execute SQL
    
    SQL = "Delete from Invdet"
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Vno= " & Val(txtVno.Text)
    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtVno.Text & " Deleted ", vbInformation
        
    EntryLast
    
    SetFocusTo SSTab1
    Exit Sub
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntrySave()
On Error GoTo errhndl
MP vbHourglass
    
    If Not ValidateControl Then
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnadd)
        Exit Sub
    End If
    
    If Not Validate() Then
        MsgBox "Please provide value for unit price. Zero price not allowed!", vbOKOnly
        SetFocusTo msfDetail
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
    ClearMsf msfDetail
    
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

    Dim mRpt As String
    Dim SpPrm(11) As String
    Dim formulas(0) As String

    mRpt = "InvFmt.rpt"
    
    SpPrm(0) = "01/01/1900"                             'From Date
    SpPrm(1) = "13/12/2090"                             'To Date
    SpPrm(2) = 0                                        'Terminal
    SpPrm(3) = 0                                        'User
    SpPrm(4) = 0
    SpPrm(5) = 0                                        'Item Id
    SpPrm(6) = 0                                        'Category Id
    SpPrm(7) = 0                                        'Size Id
    SpPrm(8) = 0                                        'Unit Id
    SpPrm(9) = "Item"                                   'Group By - Item/Category
    SpPrm(10) = txtVno.Text                             'Vno
    SpPrm(11) = 0                                       'Preview Enabled : Always Off : Used for Debug
    
    formulas(0) = "ReportTitle='Inventory Format Report'"
    
    SQL = GenReportSP("rptInventoryDetail", SpPrm)
    gCnnMst.Execute SQL

    With frmCrviewer
        .ViewReport mRpt, SpPrm(), formulas(), 0
        .Tag = "rep_Invfmt"
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
        .TableName = GetDbTable("Invtrn", gMdbMst)
        .FieldList = "Vno,doc_no,Convert(Varchar(10),rec_dat,103) as ReceiveDate"
        .CodeField = "Vno"
        .NameField = "doc_no"
        .SetFocus
        .ShowHelp
    End With
    
    txtVno.Text = Val(hlpFind.CodeText)
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

Private Sub cboItemCategory_Click()
    LoadItemsByCategogy cboItemCategory.ItemData(cboItemCategory.ListIndex), txtVno.Text
End Sub


Private Function DoShowTerminalCode() As Boolean
    
    Dim iCnt As Integer
    Dim tranType As Integer
    
    If cmbType.ListIndex = -1 Then
        Exit Function
    End If
    
    tranType = Val(cmbType.ItemData(cmbType.ListIndex))

    For iCnt = 0 To UBound(showTerminalCodeForInwardOutwardTypeList)
        If tranType = showTerminalCodeForInwardOutwardTypeList(iCnt) Then
            DoShowTerminalCode = True
            Exit Function
        End If
    Next
    
    DoShowTerminalCode = False
    Exit Function
End Function


Private Sub cmbType_LostFocus()
    If DoShowTerminalCode Then
        hlpTerminalcode.Visible = True
    Else
        hlpTerminalcode.Visible = False
    End If
End Sub

Private Sub cmdFile_Click()
    'Dim bytData() As Byte
    documentPath = ""
    On Error GoTo errhndl
    
    With filedialogue
        .Filter = gDocumentTypesFilter
        .DialogTitle = "Select Document"
        .ShowOpen
        
        If .FileTitle <> "" Then
            documentPath = .fileName
        End If
        
    End With
    
Exit Sub
errhndl:

    If Err.Number = 32755 Then
    Else
        ErrMsg
    End If
    
    Resume Next
    
End Sub

Private Sub cmdViewDocument_Click()
On Error GoTo errhndl
    
    MP vbHourglass
    
    OpenDocument lblFileName.Caption
    
    MP vbDefault
    
Exit Sub
errhndl:
    ErrMsg
    Resume Next

End Sub

Private Sub Form_Activate()
On Error GoTo errhndl
    
    MP vbHourglass
    
    SetTextBoxes
    
    MP vbDefault
    
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub


Private Sub Form_Load()

On Error GoTo errhndl
MP vbHourglass
    
    GrabActiveControl
    SetMsfDetail msfDetail
    
    zerobarredTypes = Split(gDenyZeroPriceMaterialInwardOutwardTypes, ",")
    showTerminalCodeForInwardOutwardTypeList = Split(gShowTerminalCodeForInwardOutwardTypes, ",")

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

Private Sub SetMsfDetail(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass
    
    With s_Msf
        .Cols = 6
        .Rows = 2
        .FixedRows = 1
         
        .RowHeightMin = 300
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(dColItm_id) = 800
        .ColAlignment(dColItm_id) = flexAlignLeftCenter
        
        .ColWidth(dColItm_name) = 4500
        .ColAlignment(dColItm_name) = flexAlignLeftCenter
        
        .ColWidth(dColQty) = 1000
        .ColAlignment(dColQty) = flexAlignRightCenter
        
        .ColWidth(dColUnit) = 1200
        .ColAlignment(dColUnit) = flexAlignRightCenter
        
        .ColWidth(dColrtl_rpc) = 1300
        .ColAlignment(dColrtl_rpc) = flexAlignRightCenter
        
        .ColWidth(dColAmt) = 1600
        .ColAlignment(dColAmt) = flexAlignRightCenter
        
    End With

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub SetTextBoxes()
        
    hlpFind.Visible = False
    
    FillTypeCombo
    FillCategoryCombo
    
    With hlpTerminalcode
        .SetAdoConnStr = gCnnMst
        .TableName = "TerminalConfig"
        .FieldList = "code,name,shortname"
        .CodeField = "code"
        .NameField = "ShortName"
        .SqlWhere = " actv_fg = 1"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With

    VisibleNoVisibleBtn True
    SetActiveModeNControl mEntryMode
    CenterFrmChild Me
End Sub

Private Function Validate() As Boolean
    '---Validate Zero Price
    If Not ZeroUnitPriceAllowed() Then
        Dim i As Integer
        With msfDetail
            For i = 0 To .Rows - 1
                If Val(.TextMatrix(i, dColQty)) > 0 And Val(.TextMatrix(i, dColrtl_rpc)) = 0 Then
                    Validate = False
                    Exit Function
                End If
            Next
        End With
    Else
        Validate = True
        Exit Function
    End If
    
    Validate = True
    Exit Function
   
End Function

Private Function ZeroUnitPriceAllowed() As Boolean
    
    Dim iCnt As Integer
    Dim tranType As Integer
    tranType = Val(cmbType.ItemData(cmbType.ListIndex))
    
    For iCnt = 0 To UBound(zerobarredTypes)
        If tranType = zerobarredTypes(iCnt) Then
            ZeroUnitPriceAllowed = False
            Exit Function
        End If
    Next
    
    ZeroUnitPriceAllowed = True
    Exit Function
End Function

Private Sub SaveInTmp()
On Error GoTo errhndl
Dim SQL As String
Dim i As Integer
Dim fileName As String
MP vbHourglass
    
    '''Copy file to program location
    If Trim(documentPath) <> "" Then
        fileName = CopyDocument(documentPath)
    End If
    
    gCnnMst.BeginTrans
    
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from Invtrn "
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Vno = " & Val(txtVno.Text)
        gCnnMst.Execute SQL
        
        SQL = "Delete from Invdet "
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Vno = " & Val(txtVno.Text)
        gCnnMst.Execute SQL
    End If
    
    '---Insert into Invtrn
    SQL = "Insert into  Invtrn ("
    SQL = SQL & " Vno"
    SQL = SQL & ",ter_id"
    SQL = SQL & ",export_fg"
    
    SQL = SQL & ",tran_type"
    SQL = SQL & ",doc_no"
    SQL = SQL & ",rec_dat"
    SQL = SQL & ",remarks"
    
    SQL = SQL & ",dtadat "
    SQL = SQL & ",dtatim "
    SQL = SQL & ",dtausr "
    
    SQL = SQL & ",Trng_fg"
    SQL = SQL & ",TerminalCode"
    SQL = SQL & ",[FileName]"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtVno.Text)
    SQL = SQL & "," & Val(gTerminalId)
    SQL = SQL & "," & "0"
    
    If cmbType.ListIndex <> -1 Then
        SQL = SQL & "," & Val(cmbType.ItemData(cmbType.ListIndex))
    Else
        SQL = SQL & "," & "2"   'Stock Inward
    End If
    
    SQL = SQL & "," & AQ(txtDoc_No.Text)
    SQL = SQL & "," & IIf(IsDate(mskRec_Dat.Text), ConvDatSql(mskRec_Dat.Text), "NULL")
    SQL = SQL & "," & AQ(txtRemarks.Text)
    
    SQL = SQL & "," & ConvDatSql(Date, BE_SQLSrv)
    SQL = SQL & "," & AQ(DtaTime)
    
    SQL = SQL & "," & AQ(gUser)
    
    SQL = SQL & "," & IsTrainingMode
    SQL = SQL & "," & Val(hlpTerminalcode.CodeText)
    SQL = SQL & "," & AQ(fileName)
    
    SQL = SQL & ")"
    gCnnMst.Execute SQL
    
    '---Entry in Invdet
    With msfDetail
        For i = 0 To .Rows - 1
            If Val(.TextMatrix(i, dColQty)) > 0 Then
                SQL = "Insert into  Invdet ("
                SQL = SQL & " vno "
                SQL = SQL & ",srno "
                SQL = SQL & ",itm_code "
                SQL = SQL & ",rtl_prc "
                
                SQL = SQL & ",qty "
                SQL = SQL & ",Amt "
                
                SQL = SQL & ",Trng_fg"
                
                SQL = SQL & " ) Values ("
                
                SQL = SQL & Val(txtVno.Text)
                SQL = SQL & "," & i + 1
                SQL = SQL & "," & Val(.TextMatrix(i, dColItm_id))
                SQL = SQL & "," & Val(.TextMatrix(i, dColrtl_rpc))
                
                Select Case cmbType.ItemData(cmbType.ListIndex)
                    Case 1, 2, 3, 5   ' +
                        SQL = SQL & "," & Val(.TextMatrix(i, dColQty))
                        SQL = SQL & "," & Val(.TextMatrix(i, dColAmt))
    
                    Case 11, 12, 13    ' -
                        SQL = SQL & "," & Val(.TextMatrix(i, dColQty)) * -1
                        SQL = SQL & "," & Val(.TextMatrix(i, dColAmt)) * -1
                End Select
                
                SQL = SQL & "," & IsTrainingMode
        
                SQL = SQL & ")"
                gCnnMst.Execute SQL
            End If
        Next
    End With
    
    
    gCnnMst.CommitTrans

    MsgBox "Entry No : " & txtVno.Text & " Saved ", vbInformation
    
MP vbDefault
Exit Sub
errhndl:
    gCnnMst.RollbackTrans
    ErrMsg
    
End Sub

Private Function CopyDocument(documentPath As String) As String
    Dim fileName As String
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    If Not fso.FolderExists(gDocumentPath) Then
        fso.CreateFolder (gDocumentPath)
    End If
    
    If fso.FileExists(documentPath) Then
        Dim newFilePath As String
        'Dim renameFileName As String
        fileName = txtVno.Text & "." & fso.GetExtensionName(documentPath)
        'renameFileName = txtVno.Text & "." & fso.GetExtensionName(documentPath) & "_1"
        
        newFilePath = gDocumentPath & "\" & fileName
        If fso.FileExists(newFilePath) Then
            'Call fso.MoveFile(newFilePath, gDocumentPath & "\" & fileName & "\" & renameFileName)
        End If
        
        fso.CopyFile documentPath, newFilePath, True
        
        CopyDocument = fileName
    End If
    
    Set fso = Nothing

    Exit Function
End Function

Private Sub OpenDocument(fileName)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim newFilePath As String
    newFilePath = gDocumentPath & "\" & fileName
    
    If fso.FileExists(newFilePath) Then
        Dim applicationName As String
        applicationName = getApplicationName(fso.GetExtensionName(newFilePath))
        
        Shell32Bit applicationName & " " & newFilePath
    End If
    
    Set fso = Nothing

    Exit Sub
End Sub

Private Function getApplicationName(fileExt) As String
    Dim iCnt As Integer
    Dim settings() As String
    settings = Split(gDocumentOpenIn, ",")
    
    For iCnt = 0 To UBound(settings) - 1
        Dim item() As String
        item = Split(settings(iCnt), "-")
        
        If Trim$(item(0)) = fileExt Then
            getApplicationName = item(1)
            Exit For
        End If
    Next
End Function


'Private Sub SaveDocument(documentPath As String)
'    Dim bytData() As Byte
'    On Error GoTo errhndl
'
'        If documentPath <> "" Then
'            Open documentPath For Binary As #1
'
'            ReDim bytData(FileLen(documentPath))
'        End If
'
'        Get #1, , bytData
'        Close #1
'
'    End With
'End Sub

Private Sub Nevigate(s_Mode As Nevigate)
On Error GoTo errhndl
MP vbHourglass
    
    Dim rsttmp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim iCnt As Integer
    Dim iCategory_id As Integer
    
    Select Case s_Mode
        Case MoveFirst
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " order by vno"
        Case MoveNext
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " Where 1 = 1 "
            SQL = SQL & " and vno > " & Val(txtVno.Text)
            SQL = SQL & " order by vno"
        Case MovePrev
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " Where 1 = 1 "
            SQL = SQL & " and vno < " & Val(txtVno.Text)
            SQL = SQL & " order by vno desc"
        Case MoveLast
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " order by vno desc"
        Case MoveTo
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " Where 1 = 1 "
            SQL = SQL & " and vno = " & Val(txtVno.Text)
            
    End Select
    
    OpenAdoRst rsttmp, SQL
        
        AdoRsRead rsttmp
        
        If rsttmp.RecordCount > 0 Then
            Select Case Val(rsttmp.Fields("tran_type"))
                Case 2
                    cmbType.ListIndex = 0
                Case 1
                    cmbType.ListIndex = 1
                Case 3
                    cmbType.ListIndex = 2
                Case 5
                    cmbType.ListIndex = 3
                Case 11
                    cmbType.ListIndex = 4
                Case 12
                    cmbType.ListIndex = 5
                Case 13
                    cmbType.ListIndex = 6
                Case Else
                    'do nothing
            End Select
            
            If Val(hlpTerminalcode.CodeText) > 0 Then
                hlpTerminalcode.GetNameText Val(hlpTerminalcode.CodeText)
            Else
                hlpTerminalcode.NameText = ""
            End If
            
            lblFileName.Caption = IfNullThen(rsttmp.Fields("FileName"), "")
            cmdViewDocument.Enabled = Len(Trim(lblFileName.Caption)) > 0
        End If
        
        'ReadInvdet rsttmp
        SQL = " select top 1 category_id " & vbCrLf
        SQL = SQL & " from Items" & vbCrLf
        SQL = SQL & " Inner Join Invdet on (Items.code = Invdet.itm_code)" & vbCrLf
        SQL = SQL & " Where Invdet.vno = " & Val(txtVno.Text)
        
        OpenAdoRst rs, SQL
        iCategory_id = IfNullThen(rs.Fields(0).Value, 0)
        CloseAdoRst rs
        
        For iCnt = 0 To cboItemCategory.ListCount - 1
            If Val(cboItemCategory.ItemData(iCnt)) = iCategory_id Then
                cboItemCategory.ListIndex = iCnt
                cboItemCategory.Enabled = False
                cmbType.Enabled = False
                
                Exit For
            End If
            
        Next
        
        CloseAdoRst rsttmp
    
    CalculateTotal msfDetail
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub ReadInvdet(s_rsttmp As ADODB.Recordset)
On Error GoTo errhndl
MP vbHourglass
    
    If s_rsttmp.RecordCount <= 0 Then
        MsgBox "No Records Found..", vbInformation
        SetFocusTo SSTab1
    Else
        fmeMsfdetail.Enabled = True
        
        Dim rsttmp As ADODB.Recordset
            
        SQL = "Select Invdet.Srno" & vbCrLf
        SQL = SQL & ",Invdet.Itm_code" & vbCrLf
        SQL = SQL & ",Items.ShortName " & vbCrLf
        SQL = SQL & ",Abs(Invdet.qty),units.name" & vbCrLf
        SQL = SQL & ",Invdet.rtl_prc,Abs(Invdet.amt)" & vbCrLf
        
        SQL = SQL & " From " & GetDbTable("Invdet", gMdbMst) & " Invdet " & vbCrLf
        SQL = SQL & " Left join " & GetDbTable("Items", gMdbMst) & " Items "
        SQL = SQL & " on (Invdet.itm_code=Items.code)" & vbCrLf
        
        SQL = SQL & " Left join " & GetDbTable("units", gMdbMst) & " units "
        SQL = SQL & " ON (Items.unit_id=units.code)" & vbCrLf
        
        SQL = SQL & " where Invdet.vno = " & Val(txtVno.Text) & vbCrLf
         
        OpenAdoRst rsttmp, SQL
        If rsttmp.RecordCount > 0 Then
            Set msfDetail.Recordset = rsttmp
        Else
            MsgBox "No Records Found For Detail Part", vbExclamation
            SetFocusTo SSTab1
        End If
        
        CloseAdoRst rsttmp
     End If
     
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, Shift, False
End Sub

Public Sub EnableDisable(s_Enable As Boolean)
    FmeCompanyDetail.Enabled = s_Enable
    fmeRecDetail.Enabled = s_Enable
    fmeTotals.Enabled = s_Enable
    fmeMsfdetail.Enabled = s_Enable
End Sub

Private Sub msfDetail_DblClick()
    
    Dim arr(6) As String
    Dim revVal As Variant
    Dim iEditRow As Integer
    
    With msfDetail
        iEditRow = .Row
        arr(0) = .TextMatrix(.Row, dColItm_id)
        arr(1) = .TextMatrix(.Row, dColItm_name)
        arr(2) = .TextMatrix(.Row, dColQty)
        arr(3) = .TextMatrix(.Row, dColUnit)
        arr(4) = .TextMatrix(.Row, dColrtl_rpc)
        arr(5) = .TextMatrix(.Row, dColAmt)
        arr(6) = Val(cmbType.ItemData(cmbType.ListIndex))
    End With
    
    With frmInvtrnDetails
        revVal = .Display(arr)
    End With
    
    'Update grid with new values
    With msfDetail
        .TextMatrix(iEditRow, dColQty) = revVal(dColQty)
        .TextMatrix(iEditRow, dColrtl_rpc) = revVal(dColrtl_rpc)
        .TextMatrix(iEditRow, dColAmt) = revVal(dColAmt)
    End With
    
    CalculateTotal msfDetail
    
End Sub

'Private Sub hlpItem_Validate(Cancel As Boolean)
'
'    If Val(hlpItem.CodeText) <= 0 Then
'        Cancel = True
'    Else
'        Dim rsttmp As ADODB.Recordset
'        SQL = "select UM.code,UM.[name] "
'        SQL = SQL & " from " & GetDbTable("Items", gMdbMst) & " AS RM"
'        SQL = SQL & " inner join" & GetDbTable("Units", gMdbMst) & " AS UM"
'        SQL = SQL & " ON UM.code=RM.code"
'        SQL = SQL & " where RM.code=" & Val(hlpItem.CodeText)
'
'        OpenAdoRst rsttmp, SQL, , , , gCnnMst
'        If rsttmp.RecordCount > 0 Then
'            txtUnit.Text = rsttmp.Fields("name")
'            txtUnit.Tag = rsttmp.Fields("code")
'        End If
'    End If
'End Sub

Private Sub msfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errhndl
 
   
    If LCase(lblMode.Caption) = "add" Or LCase(lblMode.Caption) = "edit" Then
        Select Case KeyCode
            Case vbKeyReturn
                msfDetail_DblClick
                
        End Select
    End If

Exit Sub
errhndl:
    Resume Next

End Sub

Private Sub txtRemarks_LostFocus()
    AskSave txtVno.Text, msfDetail, mEntryMode
End Sub


Private Sub CalculateTotal(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass
    
Dim mTotAmt As Double
Dim mRecPcs As Double

    With s_Msf
        For i = 0 To .Rows - 1
            mTotAmt = mTotAmt + Val(.TextMatrix(i, dColAmt))
            mRecPcs = mRecPcs + Val(.TextMatrix(i, dColQty))
        Next
    End With
    
    txtItemTot.Text = mTotAmt
    txtTotRecQty.Text = mRecPcs
    
MP vbDefault
Exit Sub
errhndl:
    Resume Next
    
End Sub


Private Sub ClearMsf(s_Msf As MSHFlexGrid)
    With s_Msf
        .Clear
        .Rows = 0
        .Cols = 9
    End With
End Sub



Private Sub FillTypeCombo()
    
    '   1   -   Opening Balance         -Ugadato stock
    '   2   -   Stock Inward            -Navo Stock
    '   3   -   Stock Adjustment Up     -Stock Sarbhar (Vadharo)
    '   4   -   Receive from Store      - Not used
    '   5   -   Sales Return            - Vechan Parat
    
    '   11  -   Stock Adjustment Down   -Stock Sarbhar (Gathado)
    '   12  -   Stock Waste             -Kharab Stock
    '   13  -   Issue For Sale          - Not used
    
    
    With cmbType
        .FontName = gGujaratiFontName
        .FontSize = 12
        
        .AddItem "WDz;tu Mxtuf - 1" 'Opening Stock
        .ItemData(.NewIndex) = 1
        
        .AddItem "lJtu Mxtuf - 2" 'Stock Inward
        .ItemData(.NewIndex) = 2
        
        .AddItem "Mxtuf mhCh (JDthtu) - 3" 'Stock Adjustment Up
        .ItemData(.NewIndex) = 3
    
'        .AddItem "Receive from Store - 4"
'        .ItemData(.NewIndex) = 4
    
        .AddItem "JuatK vh; (JDthtu) - 5"
        .ItemData(.NewIndex) = 5
    
        .AddItem "Mxtuf mhCh (Dxtztu) - 11" 'Stock Adjustment Down
        .ItemData(.NewIndex) = 11
        
        cmbType.AddItem "Fhtc Mxtuf - 12" 'Stock Waste
        .ItemData(.NewIndex) = 12
        
'        cmbType.AddItem "Issue for Sale - 13"
'        .ItemData(.NewIndex) = 13
        
        .ListIndex = -1
    End With
    
    
End Sub

Private Sub SSTabDetail_Click(PreviousTab As Integer)
    Select Case Val(SSTabDetail.Tab)
        Case 0
            If fmeMsfdetail.Enabled Then SetFocusTo msfDetail
        Case 1
            'do nothing
    End Select
End Sub


Private Sub FillCategoryCombo()
    SQL = "Select code,Convert(varchar(6),code) + ' - ' + shortname as shortname from Categories where actv_fg=1 order by code"
    Dim rs As ADODB.Recordset
    
    OpenAdoRst rs, SQL
    
    cboItemCategory.Font.Name = gGujaratiFontName
    cboItemCategory.Font.Size = 12
    cboItemCategory.Clear
    
    If rs.RecordCount > 0 Then
        While Not rs.EOF
        
            With cboItemCategory
                .AddItem rs.Fields("shortname").Value
                .ItemData(.NewIndex) = rs.Fields("code").Value
            End With
            
            rs.MoveNext
            
        Wend
    End If
    
    rs.Close
    Set rs = Nothing

End Sub


Private Sub LoadItemsByCategogy(s_CategoryId As Integer, s_Vno As Integer)

    SQL = "Select  " & vbCrLf
    SQL = SQL & "  code as Code" & vbCrLf
    SQL = SQL & " ,shortname as [Name]" & vbCrLf
    SQL = SQL & " ,Abs(Isnull(Invdet.qty,0)) as Qty" & vbCrLf
    SQL = SQL & " ,0 as Unit" & vbCrLf
    SQL = SQL & " ,Abs(Isnull(Invdet.rtl_prc,0)) as PricePerUnit" & vbCrLf
    SQL = SQL & " ,Abs(Isnull(Invdet.amt,0)) as Amout  " & vbCrLf
    
    SQL = SQL & " from Items " & vbCrLf
    SQL = SQL & " Left Join (Select itm_code,qty,rtl_prc,amt from Invdet  Where vno = " & s_Vno & " ) Invdet  on (Invdet.itm_code=Items.code) " & vbCrLf
    SQL = SQL & " where Items.actv_fg = 1 " & vbCrLf
    SQL = SQL & " and Items.category_id =  " & s_CategoryId & vbCrLf
    SQL = SQL & " order by Items.code"
    
    Dim rs As ADODB.Recordset
    
    OpenAdoRst rs, SQL

    If rs.RecordCount > 0 Then
        Set msfDetail.Recordset = rs
        SetGridColGujFont msfDetail, 1, 12, 1
        msfDetail.Row = 1
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub

