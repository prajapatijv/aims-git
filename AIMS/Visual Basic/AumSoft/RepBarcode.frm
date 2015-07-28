VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmRepBarcode 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   ClientHeight    =   3480
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
   ScaleHeight     =   3480
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
      TabIndex        =   6
      Top             =   2520
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
      TabIndex        =   5
      Top             =   1920
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
      Height          =   2895
      Left            =   60
      TabIndex        =   7
      Top             =   525
      Width           =   6255
      Begin VB.CheckBox chkisPaymentBarcodeLable 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Caption         =   "Print Payment Barcode label?"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   5055
      End
      Begin VB.OptionButton optSingleSideLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Caption         =   "Single Side"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton optSideBySideLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFBFC&
         Caption         =   "Side by Side"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Value           =   -1  'True
         Width           =   2175
      End
      Begin CommCtrls.ItxtBox txtLabelCount 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
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
         MaxVal          =   1000
      End
      Begin HlpN.HlpNCode hlpItem 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   661
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label Count"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   1380
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   780
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
      Caption         =   "Barcode Label"
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
      Left            =   4530
      TabIndex        =   9
      Top             =   90
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   3120
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
      Caption         =   "Barcode Label"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   2040
   End
End
Attribute VB_Name = "frmRepBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mRpt As String
Const PAYMENT_BARCODE = "00000"

Private Sub SetTextBoxes()
    
    With hlpItem
        .SetAdoConnStr = gCnnMst
        .TableName = "Items"
        .FieldList = "Code,Name,ShortName"
        .CodeField = "Code"
        .NameField = "ShortName"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
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
    
    If chkisPaymentBarcodeLable.Value = 1 Then
        GeneratePaymentBarcodeLabels
    Else
        GenerateBarcodeLabels
    End If
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_Activate()
    SetTextBoxes
End Sub

Private Sub GeneratePaymentBarcodeLabels()
    Dim SpPrm() As String
    Dim formulas() As String
    Dim mFilterText As String
    Const CONST_Category As String = "Category"
    Const CONST_Item As String = "Item"
    
    '----------------------------------------------------------------------------
    ResetReportFilters
    
    Select Case LCase(Me.Tag)
        Case LCase("rep_ItmBarcode")
            Dim intLoop As Integer
            For intLoop = 1 To Int(txtLabelCount.Text) - 1
                SetReportFilters CONST_Item, PAYMENT_BARCODE, ""
            Next intLoop

            ReDim SpPrm(0) As String
            ReDim formulas(0) As String
            
            If optSideBySideLabel.Value = True Then
                mRpt = "BarcodeLabel_Sbs.rpt"
            End If
            
            If optSingleSideLabel.Value = True Then
                mRpt = "BarcodeLabel.rpt"
            End If
    End Select

    '---Get Remarks-------------------------------------------------------------
    mFilterText = ""
    
    '----------------------------------------------------------------------------
    'formulas(1) = "ReportFilter=" & "'" & mFilterText & "'"
    'formulas(2) = "GenAt=" & "'" & ReportGenAt & "'"
    
    '----------------------------------------------------------------------------
    SQL = "Exec rptItemList " & CONST_Category & "," & CONST_Item & ",0," & Val(txtLabelCount.Text) & ",1"
    
    gCnnMst.Execute SQL

    With frmCrviewer
        .ViewReport mRpt, SpPrm(), formulas(), 0
        .Tag = "rep_BarcodeLabel"
        .Show
    End With

    '----------------------------------------------------------------------------
    ResetReportFilters

End Sub

Private Sub GenerateBarcodeLabels()
    
    Dim SpPrm() As String
    Dim formulas() As String
    Dim mFilterText As String
    Const CONST_Category As String = "Category"
    Const CONST_Item As String = "Item"
    
    '----------------------------------------------------------------------------
    ResetReportFilters
    
    Select Case LCase(Me.Tag)
        Case LCase("rep_ItmBarcode")
            Dim intLoop As Integer
            For intLoop = 1 To Int(txtLabelCount.Text) - 1
                SetReportFilters CONST_Item, Val(hlpItem.CodeText), ""
            Next intLoop

            ReDim SpPrm(0) As String
            ReDim formulas(0) As String
            
            If optSideBySideLabel.Value = True Then
                mRpt = "BarcodeLabel_Sbs.rpt"
            End If
            
            If optSingleSideLabel.Value = True Then
                mRpt = "BarcodeLabel.rpt"
            End If
            
            'formulas(0) = "ReportTitle='Barcode Label Report'"
    End Select
        
    
    '---Get Remarks-------------------------------------------------------------
    mFilterText = ""
    
    '----------------------------------------------------------------------------
    'formulas(1) = "ReportFilter=" & "'" & mFilterText & "'"
    'formulas(2) = "GenAt=" & "'" & ReportGenAt & "'"
    
    '----------------------------------------------------------------------------
    SQL = "Exec rptItemList " & CONST_Category & "," & CONST_Item & ",0," & txtLabelCount.Text
    
    gCnnMst.Execute SQL

    With frmCrviewer
        .ViewReport mRpt, SpPrm(), formulas(), 0
        .Tag = "rep_BarcodeLabel"
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
