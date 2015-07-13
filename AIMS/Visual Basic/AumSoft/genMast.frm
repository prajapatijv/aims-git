VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmGenMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "genMast.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8505
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4895
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
      TabPicture(0)   =   "genMast.frx":0442
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
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   8175
         Begin VB.CheckBox chkActv_Fg 
            Height          =   255
            Left            =   2640
            TabIndex        =   2
            Top             =   420
            Width           =   1215
         End
         Begin CommCtrls.CtxtBox txtName 
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   840
            Width           =   5295
            _ExtentX        =   9340
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
         Begin CommCtrls.ItxtBox txtCode 
            Height          =   375
            Left            =   1680
            TabIndex        =   10
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
         Begin CommCtrls.CtxtBox txtShortName 
            Height          =   375
            Left            =   1680
            TabIndex        =   4
            Top             =   1320
            Width           =   5295
            _ExtentX        =   9340
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
            Left            =   7440
            TabIndex        =   8
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Short Name"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name*"
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   907
            Width           =   630
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Code*"
            Height          =   240
            Left            =   240
            TabIndex        =   5
            Top             =   427
            Width           =   615
         End
      End
   End
   Begin HlpN.HlpNCode hlpFind 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmGenMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mEntryMode As String
Public mActCtrl As Control
Dim mTblMst As String


Public Sub EntryAdd()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "add"
    ClearScreen
    EnableDisable True
    txtCode.Text = GetMaxCode(mTblMst, True, , gCnnMst)
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
        MsgBox "No Record Selected For Edit", vbCritical
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
    
    If MsgBox("Want to Delete EntryNo " & txtCode.Text & "..???", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    mEntryMode = "delete"
    
    SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And code = " & Val(txtCode.Text)
    
    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtCode.Text & " Deleted ", vbInformation
    
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
        .TableName = GetDbTable(mTblMst, gMdbMst)
        .FieldList = "Code,Name"
        .CodeField = "Code"
        .NameField = "Name"
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
    
    VisibleNoVisibleBtn True
        
    hlpFind.Visible = False
    
    txtShortName.Font = gGujaratiFontName
    txtShortName.Font.Size = 12
    
    SetActiveModeNControl mEntryMode
    
    CenterFrmChild Me
    
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
        SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And code = " & Val(txtCode.Text)
        gCnnMst.Execute SQL
    End If
    
    SQL = "Insert into " & mTblMst & "("
    SQL = SQL & " code"
    SQL = SQL & ", name"
    SQL = SQL & ", shortName"
    SQL = SQL & ", actv_fg"
    SQL = SQL & ", dtadat"
    SQL = SQL & ", dtatim"
    SQL = SQL & ", dtausr"
    SQL = SQL & ",Trng_fg"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtCode.Text)
    SQL = SQL & "," & AQ(txtName.Text)
    SQL = SQL & "," & AQ(txtShortName.Text)
    SQL = SQL & "," & Val(chkActv_Fg.Value)
    SQL = SQL & "," & ConvDatSql(Date)
    SQL = SQL & "," & AQ(DtaTime)
    SQL = SQL & "," & AQ(gUser)
    SQL = SQL & "," & IsTrainingMode
    SQL = SQL & ")"
    
    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtCode.Text & " Saved ", vbInformation
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    'Resume Next
End Sub

Public Sub EnableDisable(s_Enable As Boolean)
   fmeMaster.Enabled = s_Enable
End Sub

Private Sub Nevigate(s_Mode As Nevigate)
On Error GoTo errhndl
MP vbHourglass
    
    Dim rsttmp As ADODB.Recordset
    
    Select Case s_Mode
        Case MoveFirst
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Order by code"
        Case MoveNext
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where code > " & Val(txtCode.Text)
        Case MovePrev
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where code < " & Val(txtCode.Text) & " order by code desc"
        Case MoveLast
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Order by code Desc"
        Case MoveTo
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where code = " & Val(txtCode.Text)
    End Select
    
    OpenAdoRst rsttmp, SQL, , , , gCnnMst
    
    With rsttmp
        If .RecordCount > 0 Then
            AdoRsRead rsttmp
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
    
    Select Case LCase(frmGenMst.Tag)
    
        Case "itmcatmst"
             mTblMst = "Categories"
             SSTab1.Caption = "Categories"
             Me.Caption = "Item Category Master"
          
        Case "locationmst"
             mTblMst = "Locations"
             SSTab1.Caption = "Locations"
             Me.Caption = "Location Master"
        
        Case "sizemst"
             mTblMst = "Sizes"
             SSTab1.Caption = "Sizes"
             Me.Caption = "Size Master"
        
        Case "unitmst"
             mTblMst = "Units"
             SSTab1.Caption = "Units"
             Me.Caption = "Unit Master"
        
    End Select
    
    Form_Resize
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub txtShortName_LostFocus()
    AskSave txtCode.Text, txtName, mEntryMode
End Sub
