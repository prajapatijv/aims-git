VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmRepInv 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   Caption         =   "7"
   ClientHeight    =   5670
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
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
      TabIndex        =   11
      Top             =   4920
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
      Left            =   6600
      TabIndex        =   10
      Top             =   4320
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
      Height          =   5055
      Left            =   60
      TabIndex        =   14
      Top             =   525
      Width           =   6255
      Begin VB.Frame Frame1 
         BackColor       =   &H00DCFBFC&
         Height          =   3255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   6015
         Begin CommCtrls.ItxtBox txtTranId 
            Height          =   375
            Left            =   1320
            TabIndex        =   9
            Top             =   2760
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
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
            AllowNull       =   -1  'True
         End
         Begin VB.ComboBox cmbGrpBy 
            Height          =   360
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2280
            Width           =   2415
         End
         Begin HlpN.HlpNCode hlpItem 
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
         End
         Begin HlpN.HlpNCode hlpCategory 
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   840
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
         End
         Begin HlpN.HlpNCode hlpSize 
            Height          =   375
            Left            =   1320
            TabIndex        =   6
            Top             =   1320
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
         End
         Begin HlpN.HlpNCode hlpUnit 
            Height          =   375
            Left            =   1320
            TabIndex        =   7
            Top             =   1800
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   661
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vno"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   2820
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Summary By"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   20
            Top             =   2340
            Width           =   1140
         End
      End
      Begin HlpN.HlpNCode hlpTerminal 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   661
      End
      Begin CommCtrls.mskDat mskFdat 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   240
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
         Left            =   4740
         TabIndex        =   1
         Top             =   240
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
      Begin HlpN.HlpNCode hlpUser 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   661
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   1267
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   787
         Width           =   795
      End
      Begin VB.Label lblTodat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4320
         TabIndex        =   13
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblFdat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   435
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
      Caption         =   "Inventory Report"
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
      Left            =   3570
      TabIndex        =   16
      Top             =   90
      Width           =   2310
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   7590
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label lblRptHeadB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Report"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   120
      Width           =   2310
   End
End
Attribute VB_Name = "frmRepInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mRpt As String

Private Sub SetTextBoxes()
    
    mskFdat.Text = GetMonthInitialDat(Date)
    msktdat.Text = Date
        
    SetForm

    With hlpTerminal
        .SetAdoConnStr = gCnnMst
        .TableName = "TerminalConfig"
        .FieldList = "Code,Name,ShortName"
        .CodeField = "code"
        .NameField = "ShortName"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
        
    With hlpUser
        .SetAdoConnStr = gCnnMst
        .TableName = "UserMast"
        .FieldList = "Uid,UName"
        .CodeField = "Uid"
        .NameField = "UName"
    End With
        
    With hlpItem
        .SetAdoConnStr = gCnnMst
        .TableName = "Items"
        .FieldList = "Code,Name,ShortName"
        .CodeField = "Code"
        .NameField = "ShortName"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
        
    With hlpCategory
        .SetAdoConnStr = gCnnMst
        .TableName = "Categories"
        .FieldList = "Code,Name,ShortName"
        .CodeField = "Code"
        .NameField = "ShortName"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
        
    With hlpSize
        .SetAdoConnStr = gCnnMst
        .TableName = "Sizes"
        .FieldList = "Code,Name,ShortName"
        .CodeField = "Code"
        .NameField = "ShortName"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
        
    With hlpUnit
        .SetAdoConnStr = gCnnMst
        .TableName = "Units"
        .FieldList = "Code,Name,ShortName"
        .CodeField = "Code"
        .NameField = "ShortName"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
        
        
    With cmbGrpBy
        .AddItem "None"
        .AddItem "TranType"
        .AddItem "Item"
        .AddItem "Category"
        .AddItem "User"
        .ListIndex = 0
    End With
        
    VisibleNoVisibleBtn False
    
    CenterFrmChild Me
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
On Error GoTo errhndl
MP vbHourglass
    
    GenerateInvenotryReport
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_Activate()
    SetTextBoxes
End Sub

Private Sub GenerateInvenotryReport()
    
    Dim SpPrm() As String
    Dim formulas() As String
    Dim mFilterText As String
    
    Select Case LCase(Me.Tag)
    
        Case LCase("rep_invdet"), LCase("rep_invsum")
            ReDim SpPrm(11) As String
            ReDim formulas(2) As String
        
            If cmbGrpBy.ListIndex = 0 Then
                mRpt = "InvDetailByVno.rpt"
            Else
                mRpt = "Inv_Summary.rpt"
            End If
            
            SpPrm(0) = mskFdat.Text                             'From Date
            SpPrm(1) = msktdat.Text                             'To Date
            SpPrm(2) = Val(hlpTerminal.CodeText)                'Terminal
            SpPrm(3) = Val(hlpUser.CodeText)                    'User
            SpPrm(4) = 0
            SpPrm(5) = Val(hlpItem.CodeText)                    'Item Id
            SpPrm(6) = Val(hlpCategory.CodeText)                'Category Id
            SpPrm(7) = Val(hlpSize.CodeText)                    'Size Id
            SpPrm(8) = Val(hlpUnit.CodeText)                    'Unit Id
            SpPrm(9) = cmbGrpBy.Text                            'Group By - Item/Category
            SpPrm(10) = IIf(Trim$(txtTranId.Text) = "", "0", Trim$(txtTranId.Text)) 'Transction Id
            SpPrm(11) = 0                                       'Preview Enabled : Always Off : Used for Debug
            
            formulas(0) = "ReportTitle='Inventory Detail Report'"
            
            SQL = GenReportSP("rptInventoryDetail", SpPrm)
            gCnnMst.Execute SQL
            
    End Select
        
    
    '---Get Remarks-------------------------------------------------------------
    mFilterText = "From Date : " & mskFdat.Text
    mFilterText = mFilterText & Space(2) & "to : " & msktdat.Text
    If Val(hlpTerminal.CodeText) > 0 Then
        mFilterText = mFilterText & Space(5) & " Terminal : " & hlpTerminal.CodeText
    End If
    If Val(hlpUser.CodeText) > 0 Then
        mFilterText = mFilterText & Space(5) & " User : " & hlpUser.NameText
    End If
'    If cmbStatus.ListIndex <> -1 Then
'        mFilterText = mFilterText & Space(5) & " Status : " & cmbStatus.Text
'    End If
    '----------------------------------------------------------------------------
    If Val(hlpItem.CodeText) > 0 Then
        mFilterText = mFilterText & Space(5) & " Item : " & hlpItem.GetFieldValue("name", Val(hlpItem.CodeText))
    End If
    If Val(hlpCategory.CodeText) > 0 Then
        mFilterText = mFilterText & Space(5) & " Category : " & hlpCategory.GetFieldValue("name", Val(hlpCategory.CodeText))
    End If
    If Val(hlpSize.CodeText) > 0 Then
        mFilterText = mFilterText & Space(5) & " Size : " & hlpSize.GetFieldValue("name", Val(hlpSize.CodeText))
    End If
    If Val(hlpUnit.CodeText) > 0 Then
        mFilterText = mFilterText & Space(5) & " Unit : " & hlpUnit.GetFieldValue("name", Val(hlpUnit.CodeText))
    End If
    If Trim$(txtTranId.Text) <> "" Then
        mFilterText = mFilterText & Space(5) & " Ticket No : " & txtTranId.Text
    End If
    '----------------------------------------------------------------------------
    formulas(1) = "ReportFilter=" & "'" & mFilterText & "'"
        
    formulas(2) = "GenAt=" & "'" & ReportGenAt & "'"
    
    
    With frmCrviewer
        .ViewReport mRpt, SpPrm(), formulas(), 0
        .Tag = "InventorySummary"
        .Show
    End With

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

Private Sub SetForm()
    
    Select Case LCase(Me.Tag)
        Case LCase("rep_invsum")
            Frame1.Visible = True
            lblRptHeadW.Caption = "Inventory Summary Report"
            lblRptHeadB.Caption = lblRptHeadW.Caption
            
        Case LCase("rep_invdet")
            Frame1.Visible = True
            lblRptHeadW.Caption = "Inventory Detail Report"
            lblRptHeadB.Caption = lblRptHeadW.Caption
            
    End Select
End Sub


