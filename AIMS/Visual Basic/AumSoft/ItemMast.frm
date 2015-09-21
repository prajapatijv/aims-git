VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmItemMast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ItemMast.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8550
   Begin TabDlg.SSTab SSTab1 
      Height          =   7300
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   12885
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
      TabCaption(0)   =   " Master >>"
      TabPicture(0)   =   "ItemMast.frx":0442
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
         Height          =   6855
         Left            =   120
         TabIndex        =   17
         Top             =   60
         Width           =   8235
         Begin VB.Frame Frame1 
            Height          =   975
            Left            =   120
            TabIndex        =   31
            Top             =   5640
            Width           =   7935
            Begin CommCtrls.NTxtBox txtMin_Qty 
               Height          =   375
               Left            =   1560
               TabIndex        =   29
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Desimal         =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxVal          =   100
               AllowNull       =   -1  'True
            End
            Begin CommCtrls.NTxtBox txtMax_Qty 
               Height          =   375
               Left            =   4620
               TabIndex        =   30
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Desimal         =   0
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
            Begin VB.Label lblMinOrderQty 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min Order Qty."
               Height          =   240
               Left            =   120
               TabIndex        =   33
               Top             =   420
               Width           =   1260
            End
            Begin VB.Label lblMaxOrderQty 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max Order Qty."
               Height          =   240
               Left            =   3000
               TabIndex        =   32
               Top             =   420
               Width           =   1320
            End
         End
         Begin VB.CheckBox chkGenBarcodeFromCode 
            Caption         =   "Generated from code"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4320
            TabIndex        =   28
            Top             =   5280
            Width           =   2775
         End
         Begin VB.CheckBox chkActv_Fg 
            Height          =   255
            Left            =   2640
            TabIndex        =   1
            Top             =   420
            Width           =   1215
         End
         Begin VB.Frame fmeValues1 
            Height          =   3255
            Left            =   150
            TabIndex        =   20
            Top             =   1800
            Width           =   7935
            Begin VB.ComboBox cmbBillFmt 
               Height          =   360
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   2640
               Width           =   3015
            End
            Begin CommCtrls.NTxtBox txtRtl_Prc 
               Height          =   375
               Left            =   1560
               TabIndex        =   7
               Top             =   1680
               Width           =   1335
               _ExtentX        =   2355
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
               AllowNull       =   -1  'True
            End
            Begin CommCtrls.NTxtBox txtDisc_Per 
               Height          =   375
               Left            =   1560
               TabIndex        =   8
               Top             =   2160
               Width           =   1335
               _ExtentX        =   2355
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
               MaxVal          =   100
               AllowNull       =   -1  'True
            End
            Begin CommCtrls.NTxtBox txtDisc_Amt 
               Height          =   375
               Left            =   4620
               TabIndex        =   9
               Top             =   2160
               Width           =   1335
               _ExtentX        =   2355
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
               AllowNull       =   -1  'True
            End
            Begin HlpN.HlpNCode hlpCategory_Id 
               Height          =   375
               Left            =   1560
               TabIndex        =   4
               Top             =   240
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   661
            End
            Begin HlpN.HlpNCode hlpSize_Id 
               Height          =   375
               Left            =   1560
               TabIndex        =   5
               Top             =   720
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   661
            End
            Begin HlpN.HlpNCode hlpUnit_id 
               Height          =   375
               Left            =   1560
               TabIndex        =   6
               Top             =   1200
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   661
            End
            Begin VB.Label lblBillFmt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bill Format*"
               Height          =   240
               Left            =   120
               TabIndex        =   25
               Top             =   2700
               Width           =   1020
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit"
               Height          =   240
               Left            =   120
               TabIndex        =   24
               Top             =   1260
               Width           =   345
            End
            Begin VB.Label lblSizeId 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Size"
               Height          =   240
               Left            =   120
               TabIndex        =   23
               Top             =   780
               Width           =   390
            End
            Begin VB.Label lblCategory_Id 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Category*"
               Height          =   240
               Left            =   120
               TabIndex        =   22
               Top             =   300
               Width           =   900
            End
            Begin VB.Label lblDisc_Amt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount Amount"
               Height          =   240
               Left            =   3000
               TabIndex        =   16
               Top             =   2220
               Width           =   1500
            End
            Begin VB.Label lblDisc_Per 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount %"
               Height          =   240
               Left            =   120
               TabIndex        =   15
               Top             =   2220
               Width           =   1005
            End
            Begin VB.Label lblRtl_Prc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Retail Price"
               Height          =   240
               Left            =   120
               TabIndex        =   14
               Top             =   1740
               Width           =   1035
            End
         End
         Begin CommCtrls.ItxtBox txtCode 
            Height          =   375
            Left            =   1680
            TabIndex        =   11
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
         Begin CommCtrls.CtxtBox txtName 
            Height          =   375
            Left            =   1680
            TabIndex        =   2
            Top             =   840
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   661
            Alignment       =   0
            MaxLength       =   60
            AutoCaps        =   0   'False
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
         Begin CommCtrls.CtxtBox txtShortName 
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   1320
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   661
            Alignment       =   0
            MaxLength       =   60
            AutoCaps        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Kanaiya-Normal"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AllowNull       =   -1  'True
         End
         Begin CommCtrls.CtxtBox txtBarcodeGenerated 
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   5220
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Locked          =   -1  'True
            Alignment       =   0
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
         Begin VB.Label lblBarcode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Bar Code*"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   5280
            Width           =   975
         End
         Begin VB.Label lblShortName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name*"
            Height          =   240
            Left            =   240
            TabIndex        =   21
            Top             =   1380
            Width           =   1140
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
            Left            =   7320
            TabIndex        =   18
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name*"
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   900
            Width           =   630
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Code*"
            Height          =   240
            Left            =   240
            TabIndex        =   12
            Top             =   420
            Width           =   615
         End
      End
   End
   Begin HlpN.HlpNCode hlpFind 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmItemMast"
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
    txtCode.Text = GetMaxCode(mTblMst, True, "Code", gCnnMst)
    txtBarcodeGenerated.Text = "0" + txtCode.Text
    chkActv_Fg.Value = 1
    SetFocusTo txtName
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryEdit(iViewMode As ViewMode)
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtCode.Text) <= 0 Then
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
    EnableDisable True
    SetFocusTo txtName

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryDelete()
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtCode.Text) <= 0 Then
        MsgBox "No Record Selected For Delete ", vbCritical
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnCancel)
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    If MsgBox("Want to Delete EntryNo " & txtCode.Text & "..???", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    mEntryMode = "delete"
    
    SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Code = " & Val(txtCode.Text)
    
    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtCode.Text & " Deleted ", vbInformation
        
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
        'SetFocusTo txtCode
        Exit Sub
    End If
    
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
        .TableName = GetDbTable(mTblMst, gMdbMst)
        .FieldList = "Code,name,shortname"
        .CodeField = "Code"
        .NameField = "name"
        .DefaultSearchCol = 1
        .SetFontParameters "", "", gGujaratiFontName, 2, 12
        
        .SetFocus
        .ShowHelp
    End With
    
    txtCode.Text = Val(hlpFind.CodeText)
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
        
    hlpFind.Visible = False
    
    With hlpCategory_Id
        .SetAdoConnStr = gCnnMst
        .TableName = "Categories"
        .FieldList = "code,name,shortname"
        .CodeField = "code"
        .NameField = "ShortName"
        .SqlWhere = " actv_fg = 1"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
    
    With hlpSize_Id
        .SetAdoConnStr = gCnnMst
        .TableName = "Sizes"
        .FieldList = "code,name,shortname"
        .CodeField = "code"
        .NameField = "ShortName"
        .SqlWhere = " actv_fg = 1"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
    
    With hlpUnit_id
        .SetAdoConnStr = gCnnMst
        .TableName = "Units"
        .FieldList = "code,name,shortname"
        .CodeField = "code"
        .NameField = "ShortName"
        .SqlWhere = " actv_fg = 1"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
    
    txtShortName.Font = gGujaratiFontName
    txtShortName.Font.Size = 12
    
    With cmbBillFmt
        .AddItem "BillFmt1 - VMUM"
        .ItemData(.NewIndex) = 1
        
        .AddItem "BillFmt2 - VP"
        .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
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
        
    gCnnMst.BeginTrans
        
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from " & GetDbTable(mTblMst, gMdbMst)
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Code = " & Val(txtCode.Text)
        gCnnMst.Execute SQL
    End If
    
    SQL = "Insert into " & mTblMst & "("
    SQL = SQL & " code"
    SQL = SQL & ",name"
    SQL = SQL & ",shortname"
    SQL = SQL & ",actv_fg"
    
    SQL = SQL & ",category_id"
    SQL = SQL & ",size_id"
    
    SQL = SQL & ",rtl_prc"
    SQL = SQL & ",disc_per"
    SQL = SQL & ",disc_amt"
    
    SQL = SQL & ", dtadat"
    SQL = SQL & ", dtatim"
    SQL = SQL & ", dtausr"
    
    SQL = SQL & ",unit_id"
    SQL = SQL & ",BillFmt"
    SQL = SQL & ",Trng_fg"
    SQL = SQL & ",min_qty"
    SQL = SQL & ",max_qty"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtCode.Text)
    SQL = SQL & "," & AQ(txtName.Text)
    SQL = SQL & "," & AQ(txtShortName.Text)
    SQL = SQL & "," & Val(chkActv_Fg.Value)
    
    SQL = SQL & "," & Val(hlpCategory_Id.CodeText)
    SQL = SQL & "," & Val(hlpSize_Id.CodeText)
    
    SQL = SQL & "," & Val(txtRtl_Prc.Text)
    SQL = SQL & "," & Val(txtDisc_Per.Text)
    SQL = SQL & "," & Val(txtDisc_Amt.Text)
    
    SQL = SQL & "," & ConvDatSql(Date)
    SQL = SQL & "," & AQ(DtaTime)
    SQL = SQL & "," & AQ(gUser)
    
    SQL = SQL & "," & Val(hlpUnit_id.CodeText)
    
    If cmbBillFmt.ListIndex = -1 Then cmbBillFmt.ListIndex = 0
    SQL = SQL & "," & Val(cmbBillFmt.ItemData(cmbBillFmt.ListIndex))
    
    SQL = SQL & "," & IsTrainingMode

    SQL = SQL & "," & Val(txtMin_Qty.Text)
    SQL = SQL & "," & Val(txtMax_Qty.Text)

    SQL = SQL & ")"
    
    gCnnMst.Execute SQL
    
    ''Save Item Barcode
    If LCase(lblMode.Caption) = "add" Then
        SQL = "Insert into ItemBarcodes ("
        SQL = SQL & "itm_code "
        SQL = SQL & ",barcode "
    
        SQL = SQL & " ) Values ("
    
        SQL = SQL & Val(txtCode.Text)
        SQL = SQL & "," & AQ(txtBarcodeGenerated.Text)
    
        SQL = SQL & ")"
    
        gCnnMst.Execute SQL
    End If
    
    gCnnMst.CommitTrans
    MsgBox "Entry No : " & txtCode.Text & " Saved ", vbInformation
    
MP vbDefault
Exit Sub
errhndl:
    gCnnMst.RollbackTrans
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
    Dim rsbarcode As ADODB.Recordset
    
    Select Case s_Mode
        Case MoveFirst
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Order by Code"
        Case MoveNext
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where Code > " & Val(txtCode.Text)
        Case MovePrev
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " where Code < " & Val(txtCode.Text) & " order by Code desc"
        Case MoveLast
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Order by Code Desc"
        Case MoveTo
            SQL = "Select top 1 * from " & GetDbTable(mTblMst, gMdbMst) & " Where Code=" & Val(txtCode.Text)
    End Select
    
    OpenAdoRst rsttmp, SQL, , , , gCnnMst
    
    With rsttmp
        If .RecordCount > 0 Then
            AdoRsRead rsttmp
            hlpCategory_Id.GetNameText Val(hlpCategory_Id.CodeText)
            hlpSize_Id.GetNameText Val(hlpSize_Id.CodeText)
            hlpUnit_id.GetNameText Val(hlpUnit_id.CodeText)
            
            If IsNull(.Fields("BillFmt").Value) Then
                cmbBillFmt.ListIndex = -1
            Else
                cmbBillFmt.ListIndex = Val(.Fields("BillFmt").Value) - 1
            End If
            
            'Read Barcode
            SQL = " select top 1 barcode " & vbCrLf
            SQL = SQL & " from ItemBarcodes" & vbCrLf
            SQL = SQL & " Inner Join Items on (Items.code = ItemBarcodes.itm_code)" & vbCrLf
            SQL = SQL & " Where Items.code = " & Val(txtCode.Text)
            
            OpenAdoRst rsbarcode, SQL
            txtBarcodeGenerated.Text = IfNullThen(rsbarcode.Fields(0).Value, "")
           
            rsbarcode.Close
            Set rsbarcode = Nothing
        
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
    
    mTblMst = "Items"
    SSTab1.Caption = "Items"

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub txtDisc_Amt_Validate(Cancel As Boolean)
    If Val(txtRtl_Prc.Text) > 0 And Val(txtDisc_Amt.Text) > 0 Then
        txtDisc_Per.Text = Round((Val(txtDisc_Amt.Text) * 100) / Val(txtRtl_Prc.Text), 2)
    End If
End Sub

Private Sub txtDisc_Per_LostFocus()
    If Val(txtDisc_Per.Text) > 0 And Val(txtRtl_Prc.Text) > 0 Then
        txtDisc_Amt.Text = Round((Val(txtRtl_Prc.Text) * Val(txtDisc_Per.Text)) / 100, 2)
    End If
End Sub

Private Sub txtMaxOrderQty_LostFocus()
    AskSave txtCode.Text, txtDisc_Amt, mEntryMode
End Sub

Private Sub txtRtl_Prc_LostFocus()
    txtDisc_Amt.MinVal = 0
    txtDisc_Amt.MaxVal = Val(txtRtl_Prc.Text)
End Sub
