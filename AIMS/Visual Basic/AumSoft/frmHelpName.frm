VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelpName 
   Caption         =   "List Of"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmHelpName.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSHFlexGrid1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmHelpName.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblName"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MSHFlexGrid2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtCode"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtName"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4683
         _Version        =   393216
         BackColorBkg    =   -2147483639
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblCode 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmHelpName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetTextBoxes()
    
    VisibleNoVisibleBtn True
    
    'SetActiveModeNControl mEntryMode
    
    CenterFrmChild Me
    
End Sub

Private Sub Form_Load()
   SetFormCaption
   SetTextBoxes
   SSTab1.Tab = 0
   'SSTab1.Tab = 1
   With MSHFlexGrid1
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .FillStyle = flexFillRepeat
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = "Code"
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "Name"
   End With
   With MSHFlexGrid2
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .FillStyle = flexFillRepeat
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = "Code"
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "Name"
   End With
End Sub

Public Sub fillgrid(rs As ADODB.Recordset)
   'If SSTab1.Tab = 0 Then
     Set MSHFlexGrid1.Recordset = rs
   'ElseIf SSTab1.Tab = 1 Then
     Set MSHFlexGrid2.Recordset = rs
   'End If
   
End Sub

'Private Sub txtCode_Change(rs As Recordset)
'    MSHFlexGrid2.Rows = 1
'    rs.Find "Code like '" & txtCode & " %' "
'    Set MSHFlexGrid2.Recordset = rs
'End Sub

Private Sub txtName_Change()
    MSHFlexGrid2.Rows = 1
    rs.Find "Name like '" & txtName & " %' "
    Set MSHFlexGrid2.Recordset = rs
End Sub
Private Sub SetFormCaption()
    Select Case LCase(frmgdnMst.Tag)
     Case "gdnmst"
          'mTblMst = "Gdnmast"
          Me.Caption = Me.Caption & "Godown Master"
          'Me.Caption = "Godown Master"
     Case "matmst"
          'mTblMst = "Matmast"
          Me.Caption = Me.Caption & "Material Master"
          'Me.Caption = "Material Master"
     Case "unitmst"
          'mTblMst = "Unitmast"
          Me.Caption = Me.Caption & "Unit Master"
          'Me.Caption = "Unit Master"
     Case "grpmst"
          'mTblMst = "Grpmast"
          Me.Caption = Me.Caption & "Group Master"
          'Me.Caption = "Group Master"
     Case "rackmst"
          'mTblMst = "Rackmast"
          Me.Caption = Me.Caption & "Rack Master"
          'Me.Caption = "Rack Master"
     Case "cbmst"
          'mTblMst = "Cbmast"
          Me.Caption = Me.Caption & "CupBoard Master"
          'Me.Caption = "CupBoard Master"
     Case "drwmst"
          'mTblMst = "Drwmast"
          Me.Caption = Me.Caption & "Drawer Master"
          'Me.Caption = "Drawer Master"
    End Select
End Sub

