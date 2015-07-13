VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmEventMast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EventMast.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8550
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6376
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
      TabPicture(0)   =   "EventMast.frx":0442
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
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   8235
         Begin VB.CheckBox chkActv_Fg 
            Height          =   255
            Left            =   2760
            TabIndex        =   2
            Top             =   780
            Width           =   1215
         End
         Begin CommCtrls.ItxtBox txtCode 
            Height          =   375
            Left            =   1800
            TabIndex        =   1
            Top             =   720
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
            TabIndex        =   3
            Top             =   1200
            Width           =   6135
            _ExtentX        =   10821
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
         Begin CommCtrls.CtxtBox txtShortName 
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   1680
            Width           =   6135
            _ExtentX        =   10821
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
         Begin HlpN.HlpNCode hlpLocation_id 
            Height          =   375
            Left            =   1800
            TabIndex        =   5
            Top             =   2160
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
         End
         Begin CommCtrls.mskDat mskFdat 
            Height          =   375
            Left            =   1800
            TabIndex        =   6
            Top             =   2640
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
         End
         Begin CommCtrls.mskDat msktdat 
            Height          =   375
            Left            =   5100
            TabIndex        =   7
            Top             =   2640
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
         End
         Begin VB.Label lblActiveEvent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ActiveEvent"
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
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Date"
            Height          =   240
            Left            =   4200
            TabIndex        =   16
            Top             =   2700
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From Date"
            Height          =   240
            Left            =   240
            TabIndex        =   15
            Top             =   2700
            Width           =   945
         End
         Begin VB.Label lblLocation_id 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location*"
            Height          =   240
            Left            =   240
            TabIndex        =   14
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label lblShortName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name*"
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   1740
            Width           =   1140
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
            Left            =   7320
            TabIndex        =   11
            Top             =   600
            Width           =   600
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name*"
            Height          =   240
            Left            =   240
            TabIndex        =   9
            Top             =   1260
            Width           =   630
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Code*"
            Height          =   240
            Left            =   240
            TabIndex        =   8
            Top             =   780
            Width           =   615
         End
      End
   End
   Begin HlpN.HlpNCode hlpFind 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmEventMast"
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
    txtCode.Text = GetMaxCode(mTblMst, True, "Code", gCnnMst)
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

    SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Code = " & Val(txtCode.Text)

    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtCode.Text & " Deleted ", vbInformation
        
    SetActiveEvent

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
        .FieldList = "Code,name,shortname"
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
    
    With hlpLocation_id
        .SetAdoConnStr = gCnnMst
        .TableName = "Locations"
        .FieldList = "code,name"
        .CodeField = "code"
        .NameField = "Name"
        .SqlWhere = " actv_fg = 1"
    End With
    
    txtShortName.Font = gGujaratiFontName
    txtShortName.Font.Size = 12
    
    VisibleNoVisibleBtn True
    
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
        
    If chkActv_Fg.Value = vbChecked Then
        SQL = "Update " & GetDbTable(mTblMst, gMdbMst)
        SQL = SQL & " Set actv_fg = 0"
        SQL = SQL & " Where 1=1"
        gCnnMst.Execute SQL
    End If
    
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Code = " & Val(txtCode.Text)
        gCnnMst.Execute SQL
    End If
    
    SQL = "Insert into " & mTblMst & "("
    SQL = SQL & " code"
    SQL = SQL & ",name"
    SQL = SQL & ",shortname"
    SQL = SQL & ",actv_fg"
    
    SQL = SQL & ",Location_id"
    SQL = SQL & ",fdat"
    SQL = SQL & ",tdat"
    
    SQL = SQL & ", dtadat"
    SQL = SQL & ", dtatim"
    SQL = SQL & ", dtausr"
    
    SQL = SQL & ",Trng_fg"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtCode.Text)
    SQL = SQL & "," & AQ(txtName.Text)
    SQL = SQL & "," & AQ(txtShortName.Text)
    SQL = SQL & "," & Val(chkActv_Fg.Value)
    
    SQL = SQL & "," & Val(hlpLocation_id.CodeText)
    SQL = SQL & "," & ConvDatSql(mskFdat.Text)
    SQL = SQL & "," & ConvDatSql(msktdat.Text)
    
    SQL = SQL & "," & ConvDatSql(Date)
    SQL = SQL & "," & AQ(DtaTime)
    SQL = SQL & "," & AQ(gUser)
    
    SQL = SQL & "," & IsTrainingMode
    
    SQL = SQL & ")"
    
    gCnnMst.Execute SQL
    
    SetActiveEvent
        
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
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Order by Code"
        Case MoveNext
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where Code > " & Val(txtCode.Text)
        Case MovePrev
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " where Code < " & Val(txtCode.Text) & " order by Code desc"
        Case MoveLast
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Order by Code Desc"
        Case MoveTo
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where Code=" & Val(txtCode.Text)
    End Select
    
    OpenAdoRst rsttmp, SQL, , , , gCnnMst
    
    With rsttmp
        If .RecordCount > 0 Then
            AdoRsRead rsttmp
            hlpLocation_id.GetNameText hlpLocation_id.CodeText
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
    
    mTblMst = "EventMast"
    SSTab1.Caption = "Events"
    
    SetActiveEvent
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub msktdat_LostFocus()
    AskSave txtCode.Text, txtName, mEntryMode
End Sub

Private Sub SetActiveEvent()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    SQL = "Select [code],[name] from EventMast where actv_fg = 1"
    OpenAdoRst rs, SQL
    
    If rs.RecordCount > 0 Then
        lblActiveEvent.Caption = "Active Event : [" & rs.Fields("code").Value & " ] " & rs.Fields("name").Value
    Else
        lblActiveEvent.Caption = "No Active Events defined."
    End If
    
    rs.Close
    Set rs = Nothing
End Sub
