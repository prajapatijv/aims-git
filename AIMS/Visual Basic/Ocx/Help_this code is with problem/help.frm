VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHlp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "help.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   10455
      Begin VB.TextBox txtAltSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         TabIndex        =   4
         Top             =   233
         Width           =   3615
      End
      Begin VB.TextBox txtFilter 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   2
         Top             =   233
         Width           =   1335
      End
      Begin VB.ComboBox cmbSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8265
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfHelp 
         Height          =   3375
         Left            =   120
         TabIndex        =   0
         Top             =   795
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5953
         _Version        =   393216
         BackColor       =   16047579
         FixedCols       =   0
         BackColorSel    =   16308668
         ForeColorSel    =   0
         BackColorBkg    =   14482428
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Records :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   2280
         TabIndex        =   3
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblFilterfld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Field"
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
         Left            =   7020
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   570
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

Private Sub msfHelp_DblClick()
    With msfHelp
        F_Code = .TextMatrix(.Row, dColCode)
        F_Name = .TextMatrix(.Row, dColName)
    End With
    Unload Me
End Sub

Public Sub FilterData()
    Dim rsttmp As New ADODB.Recordset
        
    SQL = "Select "
    If Len(f_TopN) > 0 Then
        SQL = SQL & " Top " & f_TopN & " "
    End If
    SQL = SQL & f_FieldList
    SQL = SQL & " from "
    SQL = SQL & f_TabelName
    SQL = SQL & " where 1=1 "
    SQL = SQL & " And " & f_CodeField & " >= " & Val(txtFilter.Text)
    If Len(f_SqlWhere) > 0 Then
        SQL = SQL & " And " & f_SqlWhere
    End If
    
    rsttmp.Open SQL, AdoConnStr, adOpenStatic, adLockOptimistic
    Set msfHelp.Recordset = rsttmp
    SetMsfDetail msfHelp
    
    rsttmp.Close
    Set rsttmp = Nothing
End Sub

Public Sub FilterDataAlt()
    
    Dim rsttmp As New ADODB.Recordset
    SQL = "Select "
    If Len(f_TopN) > 0 Then
        SQL = SQL & " Top " & f_TopN & " "
    End If
    SQL = SQL & f_FieldList
    SQL = SQL & " from "
    SQL = SQL & f_TabelName
    SQL = SQL & " where 1=1 "
    SQL = SQL & " And lower(left(" & f_NameField & ",len('" & txtAltSearch & "')" & ")) "
    SQL = SQL & " = " & " '" & LCase(txtAltSearch) & "'"
    If Len(f_SqlWhere) > 0 Then
        SQL = SQL & " And " & f_SqlWhere
    End If
    
    rsttmp.Open SQL, AdoConnStr, adOpenStatic, adLockOptimistic
    Set msfHelp.Recordset = rsttmp
    SetMsfDetail msfHelp
    
    rsttmp.Close
    Set rsttmp = Nothing
End Sub

Private Sub msfHelp_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyReturn
        msfHelp_DblClick
        SendKeys "{Down}"
    Case vbKeyEscape
        Unload Me
    Case vbKeySpace
        'If Len(txtFilter) = 0 Then KeyAscii = 0
    Case vbKeyBack
        KeyAscii = 0
        If Len(txtFilter) > 0 Then txtFilter = Left(txtFilter, Len(txtFilter) - 1)
        FilterData
    Case Else
        txtFilter = txtFilter & Chr(KeyAscii)
        FilterData
    End Select
End Sub

Private Sub txtAltSearch_Change()
    ConvUcase txtAltSearch
    FilterDataAlt
End Sub

Private Sub txtAltSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        txtFilter.SetFocus
    Case vbKeyDown
        msfHelp.SetFocus
    End Select
End Sub

Private Sub txtFilter_Change()
    ConvUcase txtFilter
    FilterData
End Sub

Public Sub ConvUcase(t As TextBox)
    t.SelStart = Len(t)
    t.SelLength = Len(t)
    t = UCase(t)
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        msfHelp.SetFocus
    Case vbKeyDown
        msfHelp.SetFocus
    End Select
End Sub

Private Sub SetMsfDetail(s_Msf As MSHFlexGrid)
    Dim i As Integer
    With s_Msf
        .ColWidth(0) = 1000
        For i = 1 To .Cols - 1
            .ColWidth(i) = 2500
        Next
    End With
End Sub
