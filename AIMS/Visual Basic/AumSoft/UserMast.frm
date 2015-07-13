VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmUserMast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Master"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UserMast.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6180
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
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
      TabCaption(0)   =   "User Master >>"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   5895
         Begin VB.TextBox txtPwd 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   1320
            Width           =   2295
         End
         Begin CommCtrls.ItxtBox txtUid 
            Height          =   375
            Left            =   1680
            TabIndex        =   2
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
         Begin CommCtrls.CtxtBox txtUName 
            Height          =   375
            Left            =   1680
            TabIndex        =   4
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   661
            Alignment       =   0
            MaxLength       =   15
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
            Left            =   5160
            TabIndex        =   7
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblPwd 
            AutoSize        =   -1  'True
            Caption         =   "Passward*"
            Height          =   240
            Left            =   240
            TabIndex        =   5
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label lblUName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name*"
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   900
            Width           =   1110
         End
         Begin VB.Label lblUid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Id"
            Height          =   240
            Left            =   240
            TabIndex        =   1
            Top             =   420
            Width           =   645
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
Attribute VB_Name = "frmUserMast"
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
    txtUid.Text = GetMaxCode(mTblMst, True, "Uid", gCnnMst)
    SetFocusTo txtUName
    
    lblPwd.Caption = "Passward*"
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryEdit(iViewMode As ViewMode)
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtUid.Text) <= 0 Then
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
    
    Select Case Val(gUser)
        Case Val(txtUid.Text)       'Admin
            lblPwd.Caption = "Old Passward*"
        Case Else
            lblPwd.Caption = "New Passward*"
            'Admin will Reset other's pwd
    End Select
    
    EnableDisable True
    If txtUName.Enabled Then
        SetFocusTo txtUName
    Else
        SetFocusTo txtPwd
    End If

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryDelete()
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtUid.Text) <= 0 Then
        MsgBox "No Record Selected For Delete ", vbCritical
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnCancel)
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    If MsgBox("Want to Delete EntryNo " & txtUid.Text & "..???", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    mEntryMode = "delete"
    
    SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Uid = " & Val(txtUid.Text)
    
    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtUid.Text & " Deleted ", vbInformation
        
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
        'SetFocusTo txtUid
        Exit Sub
    End If
    
    txtPwd_Validate False
    
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
        .TableName = GetDbTable("UserMast", gMdbMst)
        .FieldList = "Uid,UName"
        .CodeField = "Uid"
        .NameField = "Uname"
        .SetFocus
        .ShowHelp
    End With
    
    txtUid.Text = Val(hlpFind.CodeText)
    
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
            
    txtPwd.PasswordChar = "*"
    
    hlpFind.Visible = False
    
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
        
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Uid = " & Val(txtUid.Text)
        gCnnMst.Execute SQL
    End If
    
    SQL = "Insert into " & mTblMst & "("
    SQL = SQL & " Uid"
    SQL = SQL & ", Uname"
    SQL = SQL & ", pwd"
    SQL = SQL & ", level"
    SQL = SQL & ", status"
    
    SQL = SQL & ",Trng_fg"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtUid.Text)
    SQL = SQL & "," & AQ(txtUName.Text)
    SQL = SQL & "," & AQ(ChartoAsc(txtPwd.Text))
    SQL = SQL & ", 0"
    SQL = SQL & "," & IIf(Val(txtUid.Text) = 1001, AQ("S"), AQ("T")) 'S-Super User/T-Available
    
    SQL = SQL & "," & IsTrainingMode
    
    SQL = SQL & ")"
    
    gCnnMst.Execute SQL
            
    DefaultRights txtUid.Text
    
    MsgBox "Entry No : " & txtUid.Text & " Saved ", vbInformation
    
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
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where status in ('T','S')  Order by Uid"
        Case MoveNext
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " where status in ('T','S')  And Uid > " & Val(txtUid.Text) & " order by Uid "
        Case MovePrev
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " where status in ('T','S')  And Uid < " & Val(txtUid.Text) & " order by Uid desc"
        Case MoveLast
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where status in ('T','S')  Order by Uid Desc"
        Case MoveTo
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " where status in ('T','S')  And Uid = " & Val(txtUid.Text)
    End Select
    
    OpenAdoRst rsttmp, SQL, , , , gCnnMst
    
    With rsttmp
        If .RecordCount > 0 Then
            AdoRsRead rsttmp
            txtPwd.Text = "*********"  '"**" & IfNullThen(.Fields("Pwd").Value, "") & "***"
            
            Select Case Val(rsttmp.Fields("level").Value)
                Case enUserLevel.eAdmin, enUserLevel.eImpex, enUserLevel.eTraining
                    txtUName.Enabled = False
                Case Else
                    txtUName.Enabled = True
            End Select
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
    
    mTblMst = "UserMast"
    SSTab1.Caption = "User Master >>"
    Me.Caption = "User Master"

    Form_Resize
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub txtPwd_GotFocus()
    txtPwd.SelStart = 0
    txtPwd.SelLength = Len(txtPwd.Text)
End Sub

Private Sub txtPwd_LostFocus()
    AskSave txtUid.Text, txtUName, mEntryMode
End Sub

Private Sub txtPwd_Validate(Cancel As Boolean)
    If LCase(lblPwd.Caption) = "old passward*" Then
        Dim rst As ADODB.Recordset
        SQL = "Select 'True' from UserMast "
        SQL = SQL & " where uid = " & Val(txtUid.Text)
        SQL = SQL & " And Pwd = " & AQ(ChartoAsc(txtPwd.Text))
        
        OpenAdoRst rst, SQL, , , , gCnnMst
        With rst
            If .RecordCount <= 0 Then
                MsgBox "Wrong Passward", vbCritical
                Cancel = True
                SetFocusTo txtPwd
            Else
                lblPwd.Caption = "New Passward*"
                txtPwd.Text = ""
                Cancel = True
                SetFocusTo txtPwd
            End If
        End With
    End If
    If Len(txtPwd.Text) <= 0 Then
        Cancel = True
        SetFocusTo txtPwd
    End If
End Sub

