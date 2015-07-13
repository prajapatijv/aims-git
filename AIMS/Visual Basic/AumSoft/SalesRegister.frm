VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSalesRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Register"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   Icon            =   "SalesRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Sales Register >>"
      TabPicture(0)   =   "SalesRegister.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeValues1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPreview"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPrint"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelTicket"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Preview"
      TabPicture(1)   =   "SalesRegister.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   5535
         Begin VB.CommandButton cmdNeviate 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4800
            Picture         =   "SalesRegister.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Print"
            Top             =   3180
            Width           =   495
         End
         Begin VB.CommandButton cmdNeviate 
            BackColor       =   &H00C0C0C0&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Close"
            Top             =   3780
            Width           =   495
         End
         Begin VB.CommandButton cmdNeviate 
            Height          =   495
            Index           =   0
            Left            =   4800
            Picture         =   "SalesRegister.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "First"
            Top             =   780
            Width           =   495
         End
         Begin VB.CommandButton cmdNeviate 
            Height          =   495
            Index           =   1
            Left            =   4800
            Picture         =   "SalesRegister.frx":0BC6
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Next"
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton cmdNeviate 
            Height          =   495
            Index           =   2
            Left            =   4800
            Picture         =   "SalesRegister.frx":1008
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Previous"
            Top             =   1980
            Width           =   495
         End
         Begin VB.CommandButton cmdNeviate 
            BackColor       =   &H00E0E0E0&
            Height          =   495
            Index           =   3
            Left            =   4800
            Picture         =   "SalesRegister.frx":144A
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Last"
            Top             =   2580
            Width           =   495
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00808080&
            Height          =   7455
            Left            =   120
            ScaleHeight     =   7395
            ScaleWidth      =   4395
            TabIndex        =   10
            Top             =   240
            Width           =   4455
            Begin VB.PictureBox picPreview 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   7095
               Left            =   240
               ScaleHeight     =   7095
               ScaleMode       =   0  'User
               ScaleWidth      =   3900
               TabIndex        =   11
               Top             =   150
               Width           =   3900
            End
            Begin VB.Label Label1 
               Height          =   7215
               Left            =   120
               TabIndex        =   18
               Top             =   120
               Width           =   255
            End
         End
      End
      Begin VB.CommandButton cmdCancelTicket 
         Caption         =   "&Cancel Ticket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7170
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   7680
         Visible         =   0   'False
         Width           =   855
      End
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
         Height          =   930
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7170
         Width           =   1215
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
         Height          =   930
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7170
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7170
         Width           =   1215
      End
      Begin VB.Frame fmeValues1 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   13695
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
            Height          =   5895
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   10398
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   1
            FixedRows       =   0
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHeader 
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   873
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
   End
End
Attribute VB_Name = "frmSalesRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bPrint As Boolean

Private Enum enTkt
    dCol_tran_id = 0
    dCol_Amt = 1
    dCol_Qty = 2
    dCol_PaidAmt = 3
    dCol_DiscAmt = 4
    dCol_ChangeAmt = 5
    dCol_TerId = 6
    dCol_DtaDate = 7
    dCol_DtaTime = 8
    dCol_DtaUser = 9
    dCol_Canceled = 10
End Enum

Private Sub SetTextBoxes()
        
    SetMsfDetail
    
    LoadTicketList
    
    SSTab1.Tab = 0
    
    CenterFrmChild Me
        
End Sub

Private Sub cmdCancelTicket_Click()
    
On Error GoTo errhndl
MP vbHourglass
    
    SQL = "Update TerSaltrn "
    SQL = SQL & " Set Canceled = 1"
    SQL = SQL & " Where Tran_id = " & AQ(msfDetail.TextMatrix(msfDetail.Row, dCol_tran_id))
    
    gCnnMst.Execute SQL
    
    MsgBox "Ticket Cancelled Successfully.", vbInformation
        
    LoadTicketList
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNeviate_Click(Index As Integer)
    
    With msfDetail
        Select Case Index
            Case 0      'First
                If .Rows > 0 Then
                    .Row = 0
                End If
                
            Case 1      'Next
                If .Rows > 0 And .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
            Case 2      'Prev
                If .Rows > 0 And .Row > 0 Then
                    .Row = .Row - 1
                End If
            Case 3      'Last
                If .Rows > 0 Then
                    .Row = .Rows - 1
                End If
                        
            Case 4
                SSTab1.Tab = 0
                Exit Sub
                
            Case 5
                cmdPrint_Click
                Exit Sub
                
        End Select
    End With
    
    SSTab1_Click 1
    
End Sub

Private Sub cmdPreview_Click()
On Error GoTo errhndl
MP vbHourglass
    
    SSTab1.Tab = 1
    
    picPreview.Cls
    picPreview.CurrentX = 0
    picPreview.CurrentY = 300
    picPreview.DrawWidth = 1
    picPreview.AutoRedraw = True
    
    If msfDetail.Rows > 0 Then
        Call PrintPreview(msfDetail.TextMatrix(msfDetail.Row, dCol_tran_id), True)
    End If
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub cmdPrint_Click()
    If bPrint Then
        PrintPreview msfDetail.TextMatrix(msfDetail.Row, dCol_tran_id)
    End If
End Sub


Private Sub Command2_Click()

      
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, Shift, False
End Sub

Private Sub Form_Load()
On Error GoTo errhndl
MP vbHourglass
    
    GrabActiveControl

    CenterFrmChild Me

    Call GetPrinter(Command1)
    bPrint = gPrintEnable
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_Resize()
    CenterFormCaption Me, Me.Caption
End Sub

Private Sub SetFormCaption()
On Error GoTo errhndl
MP vbHourglass
    
    With Label1
        .BackColor = picPreview.BackColor
        .Move 120, picPreview.Top, 250, picPreview.Height
    End With

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub LoadTicketList()
    Dim i As Integer
    Dim iRow As Integer
    Dim mRow As Integer
    
    Dim rst As ADODB.Recordset
    
    SQL = "Exec stpFetchTicketList"
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    Set msfDetail.Recordset = rst
    
    SetMsfDetail
    
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub SetMsfDetail()
On Error GoTo errhndl
MP vbHourglass
    
    'Set Header Grid------------------------------------------------------------------
    With mshHeader
        .RowHeight(0) = 500
        .Cols = 11
        .ScrollBars = flexScrollBarNone
        .Font.Bold = True
        .Enabled = False
        
        .ColWidth(enTkt.dCol_tran_id) = 1700
        .ColWidth(enTkt.dCol_Amt) = 1350
        .ColWidth(enTkt.dCol_Qty) = 900
        
        .ColWidth(enTkt.dCol_PaidAmt) = 1300
        .ColWidth(enTkt.dCol_DiscAmt) = 1300
        .ColWidth(enTkt.dCol_ChangeAmt) = 1500
        
        .ColWidth(enTkt.dCol_TerId) = 900
        .ColWidth(enTkt.dCol_DtaDate) = 1300
        .ColWidth(enTkt.dCol_DtaTime) = 1100
        .ColWidth(enTkt.dCol_DtaUser) = 700
        .ColWidth(enTkt.dCol_Canceled) = 1000
        
        .TextMatrix(0, enTkt.dCol_tran_id) = "Transction Id"
        .TextMatrix(0, enTkt.dCol_Amt) = "Ticket Amount"
        .TextMatrix(0, enTkt.dCol_Qty) = "Quantity"
        
        .TextMatrix(0, enTkt.dCol_PaidAmt) = "Paid Amount"
        .TextMatrix(0, enTkt.dCol_DiscAmt) = "Disc. Amount"
        .TextMatrix(0, enTkt.dCol_ChangeAmt) = "Change Amount"
        .TextMatrix(0, enTkt.dCol_TerId) = "Terminal"
        
        .TextMatrix(0, enTkt.dCol_DtaDate) = "Date"
        .TextMatrix(0, enTkt.dCol_DtaTime) = "Time"
        .TextMatrix(0, enTkt.dCol_DtaUser) = "User"
        .TextMatrix(0, enTkt.dCol_Canceled) = "Status"
    End With
    '---------------------------------------------------------------------------------
    
    'Set Detail Grid------------------------------------------------------------------
    With msfDetail
        .FixedCols = 0
        .FixedRows = 0
        .Cols = 11
        .RowHeightMin = 350
        .Font.Bold = True
        .ScrollBars = flexScrollBarVertical
        
        .ColWidth(enTkt.dCol_tran_id) = 1700: .ColAlignment(enTkt.dCol_tran_id) = flexAlignLeftCenter
        .ColWidth(enTkt.dCol_Amt) = 1350: .ColAlignment(enTkt.dCol_Amt) = flexAlignRightCenter
        .ColWidth(enTkt.dCol_Qty) = 900: .ColAlignment(enTkt.dCol_Qty) = flexAlignRightCenter
        
        .ColWidth(enTkt.dCol_PaidAmt) = 1300: .ColAlignment(enTkt.dCol_PaidAmt) = flexAlignRightCenter
        .ColWidth(enTkt.dCol_DiscAmt) = 1300: .ColAlignment(enTkt.dCol_DiscAmt) = flexAlignRightCenter
        .ColWidth(enTkt.dCol_ChangeAmt) = 1500: .ColAlignment(enTkt.dCol_ChangeAmt) = flexAlignRightCenter
        
        .ColWidth(enTkt.dCol_TerId) = 900: .ColAlignment(enTkt.dCol_TerId) = flexAlignCenterCenter
        .ColWidth(enTkt.dCol_DtaDate) = 1300: .ColAlignment(enTkt.dCol_DtaDate) = flexAlignCenterCenter
        .ColWidth(enTkt.dCol_DtaTime) = 1100: .ColAlignment(enTkt.dCol_DtaTime) = flexAlignCenterCenter
        .ColWidth(enTkt.dCol_DtaUser) = 700: .ColAlignment(enTkt.dCol_DtaUser) = flexAlignCenterCenter
        .ColWidth(enTkt.dCol_Canceled) = 1000: .ColAlignment(enTkt.dCol_Canceled) = flexAlignCenterCenter

        .Move mshHeader.Left, mshHeader.Top + mshHeader.RowHeight(0) - 20
    End With

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

' The callback function that will monitor printer/printing status.
Private Sub Command1_Click()
        On Error Resume Next
        
    If (lpdwStatus And ASB_PRINT_SUCCESS) = ASB_PRINT_SUCCESS Or _
       (lpdwStatus And ASB_NO_RESPONSE) = ASB_NO_RESPONSE Or _
       (lpdwStatus And ASB_COVER_OPEN) = ASB_COVER_OPEN Or _
       (lpdwStatus And ASB_AUTOCUTTER_ERR) = ASB_AUTOCUTTER_ERR Or _
       ((lpdwStatus And ASB_PAPER_END_FIRST) = ASB_PAPER_END_FIRST) Or ((lpdwStatus And ASB_PAPER_END_SECOND) = ASB_PAPER_END_SECOND) Then
        isFinish = True
        status = lpdwStatus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Cancel Printer Handler
'    BiCancelStatusBack (mpHandle)
'    If Not BiCloseMonPrinter(mpHandle) = SUCCESS Then
'        MsgBox ("Failed to close printer status monitor.")
'    End If
    
    Set rstPrint = Nothing
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Tab = 1 Then
        cmdPreview_Click
    End If
        
End Sub

