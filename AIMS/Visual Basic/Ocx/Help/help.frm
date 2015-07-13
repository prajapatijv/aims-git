VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHlp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "help.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   7935
      Begin VB.TextBox txtAltSearch 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblFilter 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7935
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfHelp 
         Height          =   4335
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7646
         _Version        =   393216
         ForeColor       =   8388608
         Cols            =   1
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorSel    =   16308668
         ForeColorSel    =   8388608
         BackColorBkg    =   14482428
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
   End
End
Attribute VB_Name = "frmHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const dColCode = 0
Const dColName = 1

Dim SQL As String
Dim mFilterField As String
Dim rsttmp As New ADODB.Recordset
Dim m_Gridcols
Dim m_ColName As Integer
'

Private Sub Form_Activate()
    If msfHelp.Cols - 1 >= gDefaSearchCol Then
        msfHelp.Col = gDefaSearchCol
    End If
    msfHelp_RowColChange
End Sub

Private Sub Form_Load()
    m_Gridcols = Split(gGridCols, "~")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsttmp = Nothing
End Sub

Private Sub msfHelp_DblClick()
    If msfHelp.Rows > 1 Then
        With msfHelp
            F_Code = .TextMatrix(.Row, dColCode)
            F_Name = .TextMatrix(.Row, m_ColName)
        End With
        SendKeys "{Down}"
        Unload Me
    Else
        MsgBox "No Record Available...!!!", vbInformation
        msfHelp.SetFocus
    End If
End Sub

Public Sub FilterDataAlt()
    
    SQL = "Select "
    If Len(f_TopN) > 0 Then
        SQL = SQL & " Top " & f_TopN & " "
    End If
    SQL = SQL & f_FieldList
    SQL = SQL & " from "
    SQL = SQL & f_TabelName
    SQL = SQL & " where 1=1 "
    SQL = SQL & " And " & f_NameField & " Like " & "'" & txtAltSearch.Text & "%'"
    If Len(f_SqlWhere) > 0 Then
        SQL = SQL & " And " & f_SqlWhere
    End If
    
    rsttmp.Open SQL, gAdoConnStr, adOpenStatic, adLockOptimistic
    
    If rsttmp.RecordCount > 0 Then
        txtAltSearch.BackColor = &H8000000F
        Set msfHelp.Recordset = rsttmp
        SetMsfDetail msfHelp
    Else
        txtAltSearch.BackColor = &H8080FF
    End If
    
    rsttmp.Close

End Sub

Private Sub msfHelp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            msfHelp_DblClick
            SendKeys "{Down}"
        Case vbKeyEscape
            Unload Me
        Case vbKeyTab
            KeyAscii = 0
        Case vbKeyBack
            KeyAscii = 0
            If Len(txtAltSearch) > 0 Then txtAltSearch = Left(txtAltSearch, Len(txtAltSearch) - 1)
            FilterDataAlt
        Case Else
            txtAltSearch.Text = txtAltSearch.Text & Chr(KeyAscii)
            FilterDataAlt
    End Select
End Sub

Private Sub msfHelp_RowColChange()

    Dim iCols As Integer
    Dim mCol As Integer
    
    For iCols = 0 To UBound(m_Gridcols)
        mCol = m_Gridcols(iCols)
        If msfHelp.Col = mCol Then
            txtAltSearch.Font.Name = gGridFontName
            txtAltSearch.Font.Size = gGridFontSize
            Exit For
        Else
            txtAltSearch.Font.Name = "Arial"
            txtAltSearch.Font.Size = 10
        End If
    Next
    
    lblFilter.Caption = StrConv(msfHelp.TextMatrix(0, msfHelp.Col), vbProperCase) & " : "
    HlpN.f_NameField = msfHelp.TextMatrix(0, msfHelp.Col)
    
End Sub

Private Sub txtAltSearch_Change()
    FilterDataAlt
End Sub

Private Sub txtAltSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            msfHelp.SetFocus
    End Select
End Sub

Public Sub SetMsfDetail(s_Msf As MSHFlexGrid)
    Dim i As Integer
    With s_Msf
    
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .Font.Size = 11
        .RowHeightMin = 350
        
        .ColWidth(0) = 900
        .TextMatrix(0, 0) = StrConv(.TextMatrix(0, 0), vbProperCase)
        
        For i = 1 To .Cols - 1
            .TextMatrix(0, i) = StrConv(.TextMatrix(0, i), vbProperCase)
            .ColWidth(i) = 3000
            
            If Trim$(LCase(.TextMatrix(0, i)) = Trim$(LCase(f_NameField))) Then
                m_ColName = i
            End If
        Next
        
        SetGridFont s_Msf
    End With
End Sub

Private Sub SetGridFont(s_Msf As MSHFlexGrid)

    Dim iRow As Integer
    Dim iCols As Integer
    Dim mCol As Integer
    Dim mColPrev As Integer
    
    mColPrev = msfHelp.Col
    
    For iCols = 0 To UBound(m_Gridcols)
        
        mCol = m_Gridcols(iCols)
        
        With s_Msf
            .Redraw = False
            If .Rows <> 0 Then
                For iRow = 1 To .Rows - 1
                    .Row = iRow
                    
                    .Col = mCol
                    .CellFontName = gGridFontName
                    .CellFontSize = gGridFontSize
                Next
            
                .Row = 1
                .Col = mColPrev
            End If
            .Redraw = True
        End With
        
    Next
    
End Sub
