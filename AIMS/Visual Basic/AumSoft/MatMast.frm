VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#2.0#0"; "CommCtrls.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMatMast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Master"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7095
   Begin VB.Frame fmeMatMast 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin CommCtrls.CtxtBox txtName 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         Alignment       =   0
      End
      Begin CommCtrls.CtxtBox txtAliasName 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Alignment       =   0
      End
      Begin CommCtrls.CtxtBox txtCode 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   795
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Alignment       =   0
      End
      Begin VB.Label lblAlsName 
         Caption         =   "Short Name"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   555
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   855
         Width           =   495
      End
   End
   Begin VB.Frame fmeMsfList 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   2685
      Width           =   7095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfList 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
   End
End
Attribute VB_Name = "frmMatMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mEntryMode As String
Public mActCtrl As Control
Public tmp_tMast As String
Public tMast As String

Public Sub EntryAdd()
    mEntryMode = "add"
    ClearScreen
    
    EnableDisable True
    
    txtCode.Text = GetMaxCode(tMast)
    SetFocusTo txtName
    
End Sub

Public Sub EntryEdit()
    mEntryMode = "edit"
    EnableDisable True
    SetFocusTo txtName
End Sub

Public Sub EntryDelete()
On Error GoTo ErrHndl
    
    mEntryMode = "delete"
    
    SQL = "Delete from AcMast"
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And code = " & Val(txtCode.Text)
    
    gCnnInv.Execute SQL
    
    MsgBx "Entry No : " & txtCode.Text & " Deleted ", jvOkOnly
        
    Exit Sub
    
ErrHndl:
    MsgBx "Entry not deleted due to Error"
End Sub

Public Sub EntrySave()
    mEntryMode = "save"
    SaveInTmp
    EnableDisable False
End Sub

Public Sub EntrySaveNAdd()
    mEntryMode = "savenadd"
    EntrySave
    EntryAdd
End Sub

Public Sub EntryCancel()
    mEntryMode = "cancel"
    ClearScreen
    EnableDisable False
End Sub

Public Sub EntryPrint()
    mEntryMode = "print"

End Sub

Public Sub EntryFirst()
    mEntryMode = "first"
    EnableDisable False
    Nevigate MoveFirst
End Sub

Public Sub EntryNext()
    mEntryMode = "next"
    EnableDisable False
    Nevigate MoveNext
End Sub

Public Sub EntryPrev()
    mEntryMode = "prev"
    EnableDisable False
    Nevigate MovePrev
End Sub

Public Sub EntryLast()
    mEntryMode = "last"
    EnableDisable False
    Nevigate MoveLast
End Sub

Public Sub EntryExit()
    Unload Me
End Sub
Private Sub SetTextBoxes()
    VisibleNoVisibleBtn True
    SetActiveModeNControl mEntryMode
End Sub

Private Sub Form_Activate()
    SetTextBoxes
End Sub

Private Sub Form_Load()
   If FormSelected = "Group" Then
      'tmp_tMast = "GrpMast"
      tMast = "GrpMast"
      frmMatMast.Caption = "GROUP MASTER"
   ElseIf FormSelected = "Material" Then
      'tmp_tMast = "MaterialMast"
      tMast = "MaterialMast"
      frmMatMast.Caption = "MATERIAL MASTER"
   ElseIf FormSelected = "Unit" Then
      'tmp_tMast = "UnitMast"
      tMast = "UnitMast"
      frmMatMast.Caption = "UNIT MASTER"
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mEntryMode = ""
    Set mActCtrl = Nothing
    VisibleNoVisibleBtn False
End Sub

Private Sub SaveInTmp()

    If AdoIsTable("tmp_tMast", gCnnInv) Then
        gCnnInv.Execute "drop table tmp_tMast"
    End If
        
    If FormSelected = "Group" Then
       gCnnInv.Execute SqlGrpMast("tmp_tMast")
    ElseIf FormSelected = "Material" Then
        gCnnInv.Execute SqlMaterialMast("tmp_tMast")
    ElseIf FormSelected = "Unit" Then
       gCnnInv.Execute SqlUnitMast("tmp_tMast")
    End If
    
    
    
    
    SQL = "Insert into " & tMast & "("
    SQL = SQL & " code"
    SQL = SQL & ",alsname"
    SQL = SQL & ",name"
    
''    SQL = SQL & ",grp"
''    SQL = SQL & ",grp1"
''    SQL = SQL & ",grp2"
''
''    SQL = SQL & ",add1"
''    SQL = SQL & ",add2"
''    SQL = SQL & ",pncd"
''    SQL = SQL & ",city"
''    SQL = SQL & ",Dist"
''    SQL = SQL & ",State"
''
''    SQL = SQL & ",phone"
''    'SQL = SQL & ",phone1"
''    'SQL = SQL & ",phone2"
''    SQL = SQL & ",mobile"
''
''    'SQL = SQL & ",crdays"
''    SQL = SQL & ",cstno"
''    SQL = SQL & ",gstno"
''    SQL = SQL & ",panno"
''    SQL = SQL & ",email"
''    SQL = SQL & ",www"
    
    SQL = SQL & ",Status"
    SQL = SQL & ",dta_dat"
    SQL = SQL & ",dta_user"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtCode.Text)
    SQL = SQL & "," & AQ(txtName.Text)
    SQL = SQL & "," & AQ(txtAliasName.Text)
''    SQL = SQL & "," & Val(hlpGrp.GetCode)
''    SQL = SQL & "," & Val(hlpGrp1.GetCode)
''    SQL = SQL & "," & Val(hlpGrp2.GetCode)
''
''    SQL = SQL & "," & AQ(txtAdd1.Text)
''    SQL = SQL & "," & AQ(txtAdd2.Text)
''    SQL = SQL & "," & AQ(txtPinCd.Text)
''    SQL = SQL & "," & AQ(txtCity.Text)
''    SQL = SQL & "," & AQ(txtDist.Text)
''    SQL = SQL & "," & AQ(txtState.Text)
''
''    SQL = SQL & "," & AQ(txtPhoneNo.Text)
''    'SQL = SQL & "," & AQ(txtPhoneNo1.Text)
''    'SQL = SQL & "," & AQ(txtPhoneNo.Text)
''    SQL = SQL & "," & AQ(txtMobNo.Text)
''
''    'SQL = SQL & "," & val(txtPhoneNo.Text)
''    SQL = SQL & "," & AQ(txtCstNo.Text)
''    SQL = SQL & "," & AQ(txtGstNo.Text)
''    SQL = SQL & "," & AQ(txtPanNo.Text)
''    SQL = SQL & "," & AQ(txtEmail.Text)
''    SQL = SQL & "," & AQ(txtWWW.Text)
    
    SQL = SQL & "," & AQ("T")
    'SQL = SQL & "," & allow comp
    SQL = SQL & "," & ConvDatSql(Date, BE_SQLSrv)
    SQL = SQL & "," & AQ(gUser)
    
    SQL = SQL & ")"
    
    gCnnInv.Execute SQL
    
End Sub

Private Sub Nevigate(s_Mode As Nevigate)
    
    Dim rsttmp As ADODB.Recordset
    
    Select Case s_Mode
        Case MoveFirst
            SQL = "Select top 1 * from " & tMast & " Where status = " & "'T'" & " Order by code"
        Case MoveNext
            SQL = "Select top 1 * from " & tMast & " where status =" & "'T'" & " And code > " & Val(txtCode.Text)
        Case MovePrev
            SQL = "Select top 1 * from " & tMast & " where status =" & "'T'" & " And code < " & Val(txtCode.Text)
        Case MoveLast
            SQL = "Select top 1 * from " & tMast & " Where status =" & "'T'" & " Order by code Desc"
    End Select
    
    OpenAdoRst rsttmp, SQL
    
    With rsttmp
        If .RecordCount > 0 Then
            AdoRsRead rsttmp
            'txtCode.Text = .Fields("code").Value
            'txtName.Text = .Fields("name").Value
        End If
    End With
    
    rsttmp.Close
    Set rsttmp = Nothing
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, Shift
End Sub

'Private Sub Form_Load()
'    GrabActiveControl
'End Sub

Public Sub EnableDisable(s_Enable As Boolean)
    'fmeAcDetails.Enabled = s_Enable
    'fmePersonalDetails.Enabled = s_Enable
    'fmeBusinessDetails.Enabled = s_Enable
End Sub


