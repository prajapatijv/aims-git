VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCompSel 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame fmeReport 
      BackColor       =   &H00DCFBFC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   60
      TabIndex        =   3
      Top             =   525
      Width           =   6255
      Begin VB.CheckBox chkVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Caption         =   "   Donot Show EveryTime When Program Starts"
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
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3360
         Width           =   5295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
         Height          =   3135
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5530
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   4320
      Left            =   0
      Top             =   0
      Width           =   7770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Foundry Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2370
      TabIndex        =   5
      Top             =   90
      Width           =   4185
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   7590
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label lblRepHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Foundry Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   4185
   End
End
Attribute VB_Name = "frmCompSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Const dColCheck = 0
Const dColName = 1
Const dColYear = 2
Const dColCode = 3
Const dColInvDb = 4
Const dColRepPath = 5
Const dColResPath = 6
Const dColDefaConum = 7

Private Sub SetTextBoxes()
    
    VisibleNoVisibleBtn False
    
    CenterFrmChild Me
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOk_Click()
    
    Dim i As Integer
    
    SQL = "Update " & GetDbTable("CompMast", gMdbMst)
    SQL = SQL & " set Visible = " & IIf(chkVisible.Value = vbChecked, AQ("F"), AQ("T"))
    gCnnMst.Execute SQL
            
    With msfDetail
        For i = 1 To .Rows - 1
            SQL = "Update " & GetDbTable("CompMast", gMdbMst)
            SQL = SQL & " set DefaConum = " & IIf(.TextMatrix(i, dColCheck) = "a", AQ("T"), AQ("F"))
            SQL = SQL & " Where Code = " & Val(.TextMatrix(i, dColCode))
            SQL = SQL & " And FYear = " & AQ(.TextMatrix(i, dColYear))
            gCnnMst.Execute SQL
        Next
    End With
        
    If SetDefaConum(msfDetail.TextMatrix(msfDetail.Row, dColCode)) Then
        
'        CreatePkgTrnTables
'        AddPkgConstraintsTrn
'
'        If InStr(1, Command$, "install=check", vbTextCompare) > 0 Then
'            chkStructure
'            FormCollectionDetaults
'        End If
'
'        With msfDetail
'            SetStatBarPanels StrConv(.TextMatrix(.Row, dColCode) & " - " & .TextMatrix(.Row, dColName) & " , " & .TextMatrix(.Row, dColYear), vbProperCase), gUser & " - " & gUserName
'        End With
'        Unload Me
    End If
            
End Sub
    
Private Sub Form_Activate()
    SetTextBoxes
    
    ReadCompMast
    
    CenterFrmNonChild Me
    
End Sub

Private Sub ReadCompMast()
    
    Dim rst As ADODB.Recordset
    
    SQL = "Select [Name] as CompanyName"
    SQL = SQL & ",FYear as Year"
    SQL = SQL & ",Code"
    SQL = SQL & ",InvDb"
    SQL = SQL & ",RepPath"
    SQL = SQL & ",ResPath"
    SQL = SQL & ",DefaConum"
    SQL = SQL & " From " & GetDbTable("Compmast", gMdbMst)
    SQL = SQL & " Order by Code,Fyear"
    
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    With rst
        If rst.RecordCount > 0 Then
            Set msfDetail.Recordset = rst
            
            SetMsfDetail msfDetail
            
            With msfDetail
                If .Rows > 1 Then .Row = 1
                .Col = 1
                .ColSel = .Cols - 1
            End With
            
            With msfDetail
                For i = 1 To .Rows - 1
                    If LCase(.TextMatrix(i, dColDefaConum)) = "t" Then
                        .Row = i
                        msfDetail_DblClick
                    End If
                Next
            End With
        Else
            Unload frmCompSel
            frmCompMast.Tag = "FirstConum"
            frmCompMast.Show
        End If
    End With
End Sub

Private Sub SetMsfDetail(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass
    
    With s_Msf
        .FixedCols = 1
        .RowHeightMin = 350
        .RowHeight(0) = 400
        .Font.Bold = True
        .ScrollBars = flexScrollBarVertical
        
        .ColAlignment(0) = flexAlignCenterCenter
            
        .ColWidth(dColCheck) = 400
        .ColAlignment(dColCheck) = flexAlignLeftCenter
            
        .ColWidth(dColName) = 4540 + 1050
        .ColAlignment(dColName) = flexAlignLeftCenter
        
        .ColWidth(dColYear) = 0 ''1050
        .ColAlignment(dColYear) = flexAlignCenterCenter
        
        .ColWidth(dColCode) = 0
        .ColWidth(dColInvDb) = 0
        .ColWidth(dColRepPath) = 0
        .ColWidth(dColResPath) = 0
        .ColWidth(dColDefaConum) = 0
        
    End With

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub msfDetail_DblClick()

    Dim iRow As Integer
    Dim mRow As Integer
    
    With msfDetail
        mRow = .Row
        For iRow = 1 To .Rows - 1
            .Col = dColCheck
            .CellFontName = "Marlett"

            If mRow = iRow Then
                .TextMatrix(iRow, dColCheck) = "a"
            Else
                .TextMatrix(iRow, dColCheck) = ""
            End If
        Next
        
        .Col = 1
        .ColSel = .Cols - 1
    End With
        
        
End Sub

Private Function SetDefaConum(f_Conum As Integer) As Boolean
    
    SetDefaConum = False
    
    Dim rst As ADODB.Recordset
    
    SQL = "Select code,Invdb,RepPath,ResPath "
    SQL = SQL & " From CompMast "
    SQL = SQL & " Where code =" & f_Conum
    
    OpenAdoRst rst, SQL, , , , gCnnMst
    With rst
        If .RecordCount > 0 Then
            
            gPathReport = .Fields("RepPath").Value
            gPathResources = .Fields("ResPath").Value
            gDefaComp = .Fields("Code").Value
            
            If Len(gPathReport) = 0 Or Len(gPathResources) = 0 Then
                MsgBox "Package Closed Due to Insufficent Company Information...", vbCritical
                End
            End If
            
            SetDefaConum = True
        Else
            MsgBox "No Company Found", vbCritical
            'End
        End If
    End With
    
    CloseAdoRst rst
    
End Function

Private Sub SetStatBarPanels(s_Conum As String, s_User As String)
    With mdiMainMenu.StatBar
        .Panels(3).Text = ">>> " & s_Conum
        .Panels(4).Text = ">>> " & s_User
    End With
End Sub

