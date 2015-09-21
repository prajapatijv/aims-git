VERSION 5.00
Begin VB.Form frmRepItem 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12900
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
   ScaleHeight     =   5025
   ScaleWidth      =   12900
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
      Left            =   11520
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   11520
      TabIndex        =   0
      Top             =   3600
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
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   525
      Width           =   11055
      Begin VB.Frame Frame1 
         BackColor       =   &H00DCFBFC&
         Height          =   4095
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   10695
         Begin VB.CommandButton cmdMoveLeft 
            Caption         =   "<"
            Height          =   495
            Left            =   5040
            TabIndex        =   14
            Top             =   2400
            Width           =   495
         End
         Begin VB.CommandButton cmdMiveRight 
            Caption         =   ">"
            Height          =   495
            Left            =   5040
            TabIndex        =   13
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton cmdMoveLeftAll 
            Caption         =   "<<"
            Height          =   495
            Left            =   5040
            TabIndex        =   12
            Top             =   3000
            Width           =   495
         End
         Begin VB.CommandButton cmdMoveRightAll 
            Caption         =   ">>"
            Height          =   495
            Left            =   5040
            TabIndex        =   11
            Top             =   1200
            Width           =   495
         End
         Begin VB.ListBox lstSelected 
            BackColor       =   &H00FFFFFF&
            Height          =   2940
            Left            =   5880
            TabIndex        =   9
            Top             =   960
            Width           =   4695
         End
         Begin VB.ListBox lstAll 
            BackColor       =   &H00FFFFFF&
            Height          =   2940
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   4695
         End
         Begin VB.OptionButton optItem 
            BackColor       =   &H00DCFBFC&
            Caption         =   "Item"
            Height          =   240
            Left            =   1680
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optCategory 
            BackColor       =   &H00DCFBFC&
            Caption         =   "Category"
            Height          =   240
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Records"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   6000
            TabIndex        =   16
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label lblAll 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "All Records"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   1050
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   1267
         Width           =   375
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   360
      Left            =   0
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblRptHeadW 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Report"
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
      Left            =   9690
      TabIndex        =   4
      Top             =   90
      Width           =   1620
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   11280
      X2              =   11280
      Y1              =   120
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   12480
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label lblRptHeadB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Report"
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
      Left            =   9720
      TabIndex        =   3
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmRepItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mRpt As String

Private Sub SetTextBoxes()
    
        
    lstAll.Font.Name = gGujaratiFontName
    lstSelected.Font.Name = gGujaratiFontName
    lstAll.Font.Size = 12
    lstSelected.Font.Size = 12
    
    optItem.Value = True
    FillListBoxes "item"
    
    VisibleNoVisibleBtn False
    
    CenterFrmChild Me
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMiveRight_Click()
    MoveListItems "left2right"
End Sub

Private Sub cmdMoveLeft_Click()
    MoveListItems "right2left"
End Sub

Private Sub cmdMoveLeftAll_Click()
    Dim iCnt As Integer
    For iCnt = lstSelected.ListCount - 1 To 0 Step -1
        lstSelected.Selected(iCnt) = True
        MoveListItems ("right2left")
    Next
End Sub

Private Sub cmdMoveRightAll_Click()
    Dim iCnt As Integer
    For iCnt = lstAll.ListCount - 1 To 0 Step -1
        lstAll.Selected(iCnt) = True
        MoveListItems ("left2right")
    Next
End Sub

Public Sub cmdPrint_Click()
On Error GoTo errhndl
MP vbHourglass
    
    GenerateItemReport
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_Activate()
    SetTextBoxes
End Sub

Private Sub GenerateItemReport()
    
    Dim SpPrm() As String
    Dim formulas() As String
    Dim mFilterText As String
    Const CONST_Category As String = "Category"
    Const CONST_Item As String = "Item"
    
    '----------------------------------------------------------------------------
    ResetReportFilters
    
    Select Case LCase(Me.Tag)
    
        Case LCase("rep_ItemList")
            ReDim SpPrm(0) As String
            ReDim formulas(2) As String
            
            mRpt = "ItemList.rpt"
            
            formulas(0) = "ReportTitle='Item List Report'"
            
        Case LCase("rep_ItmLst_MinMaxOrderQty")
            ReDim SpPrm(0) As String
            ReDim formulas(2) As String
            
            mRpt = "ItemListMinMaxOrderQty.rpt"
            
            formulas(0) = "ReportTitle='Item Min/Max Qty Report'"

    End Select
        
    
    '---Get Remarks-------------------------------------------------------------
    mFilterText = ""
    
    '----------------------------------------------------------------------------
    Dim iCnt As Integer
    Dim strList As String
    
    '-- Set Category Filter
    For iCnt = 0 To lstSelected.ListCount - 1
        SetReportFilters IIf(optCategory.Value = True, CONST_Category, CONST_Item), lstSelected.ItemData(iCnt), ""
        strList = strList & lstSelected.List(iCnt) & ","
    Next
    strList = Trim$(strList)
    If Len(strList) > 0 Then
        mFilterText = mFilterText & Space(5) & IIf(optCategory.Value = True, " By Category  ", " By Item ")
    End If
    
    
    '----------------------------------------------------------------------------
    formulas(1) = "ReportFilter=" & "'" & mFilterText & "'"
        
    formulas(2) = "GenAt=" & "'" & ReportGenAt & "'"
    
    
    '----------------------------------------------------------------------------
    SQL = "Exec rptItemList " & CONST_Category & "," & CONST_Item & ",0"
    
    gCnnMst.Execute SQL

    With frmCrviewer
        .ViewReport mRpt, SpPrm(), formulas(), 0
        .Tag = mRpt
        .Show
    End With

    '----------------------------------------------------------------------------
    ResetReportFilters

End Sub

Private Sub Form_Resize()
    
    With Shape1
        .BorderWidth = 3
        .Move 0, 0, Me.Width, Me.Height
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleNoVisibleBtn False, True
End Sub


Private Sub FillListBoxes(s_Mode As String)

    Dim rs As ADODB.Recordset
    
    If LCase(s_Mode) = "item" Then
        SQL = "Select code,convert(varchar(6),code)+' - '+shortname as shortname from Items"
    Else
        SQL = "Select code,convert(varchar(6),code)+' - '+shortname shortname from Categories"
    End If
    
    OpenAdoRst rs, SQL
    
    lstAll.Clear
    lstSelected.Clear
    
    
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            lstAll.AddItem (rs.Fields("shortname").Value)
            lstAll.ItemData(lstAll.NewIndex) = rs.Fields("code").Value
            rs.MoveNext
        Wend
    End If
    
    
End Sub


Private Sub lstAll_DblClick()
    MoveListItems "left2right"
End Sub

Private Sub lstSelected_DblClick()
    MoveListItems "right2left"
End Sub

Private Sub optCategory_Click()
    FillListBoxes "category"
End Sub

Private Sub optItem_Click()
    FillListBoxes "item"
End Sub

Private Sub MoveListItems(s_Mode As String)

    Select Case LCase(s_Mode)
        Case "left2right"
            If lstAll.ListIndex <> -1 Then
                lstSelected.AddItem lstAll.Text
                lstSelected.ItemData(lstSelected.NewIndex) = lstAll.ItemData(lstAll.ListIndex)
                lstAll.RemoveItem (lstAll.ListIndex)
            End If
            
        Case "right2left"
            If lstSelected.ListIndex <> -1 Then
                lstAll.AddItem lstSelected.Text
                lstAll.ItemData(lstAll.NewIndex) = lstSelected.ItemData(lstSelected.ListIndex)
                lstSelected.RemoveItem (lstSelected.ListIndex)
            End If
            
    End Select
    
End Sub
