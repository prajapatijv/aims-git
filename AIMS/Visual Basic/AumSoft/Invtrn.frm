VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.2#0"; "MyHelp.ocx"
Begin VB.Form frmInvtrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Inward/outward Entry"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2
   Icon            =   "Invtrn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   12300
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   12726
      _Version        =   393216
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
      TabCaption(0)   =   "Inward/Outward Entry"
      TabPicture(0)   =   "Invtrn.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTabDetail"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmeTotals"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fmeRecDetail"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FmeCompanyDetail"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   30
         Top             =   -600
         Width           =   1455
      End
      Begin VB.Frame FmeCompanyDetail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   330
         Width           =   12015
         Begin CommCtrls.ItxtBox txtVno 
            Height          =   375
            Left            =   10800
            TabIndex        =   1
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   240
            TabIndex        =   38
            Top             =   300
            Width           =   600
         End
         Begin VB.Label lblVno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vno*"
            Height          =   240
            Left            =   10200
            TabIndex        =   14
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame fmeRecDetail 
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   945
         Width           =   12015
         Begin VB.ComboBox cmbType 
            Height          =   360
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   247
            Width           =   3255
         End
         Begin CommCtrls.mskDat mskRec_Dat 
            Height          =   375
            Left            =   10800
            TabIndex        =   4
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
            AllowNull       =   -1  'True
         End
         Begin CommCtrls.CtxtBox txtDoc_No 
            Height          =   375
            Left            =   5520
            TabIndex        =   3
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Alignment       =   0
            MaxLength       =   30
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
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   240
            Left            =   240
            TabIndex        =   37
            Top             =   307
            Width           =   480
         End
         Begin VB.Label lblReqNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DocNo"
            Height          =   240
            Left            =   4680
            TabIndex        =   15
            Top             =   307
            Width           =   645
         End
         Begin VB.Label lblReqDat 
            AutoSize        =   -1  'True
            Caption         =   "Rec. Date"
            Height          =   240
            Left            =   9480
            TabIndex        =   16
            Top             =   300
            Width           =   900
         End
      End
      Begin VB.Frame fmeTotals 
         Height          =   1245
         Left            =   135
         TabIndex        =   32
         Top             =   5820
         Width           =   12045
         Begin CommCtrls.CtxtBox txtRemarks 
            Height          =   855
            Left            =   1320
            TabIndex        =   13
            Top             =   195
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1508
            BackColor       =   14482428
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
         Begin CommCtrls.NTxtBox txtTotRecQty 
            Height          =   375
            Left            =   8640
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   195
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Locked          =   -1  'True
            BackColor       =   14482428
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
         Begin CommCtrls.NTxtBox txtItemTot 
            Height          =   375
            Left            =   8640
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   675
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Locked          =   -1  'True
            BackColor       =   14482428
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
         Begin VB.Label lblRemarks 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblItemTot 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            Height          =   240
            Left            =   7320
            TabIndex        =   27
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label lblTotRecQty 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Qty"
            Height          =   240
            Left            =   7320
            TabIndex        =   25
            Top             =   240
            Width           =   810
         End
      End
      Begin TabDlg.SSTab SSTabDetail 
         Height          =   3855
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   12045
         _ExtentX        =   21246
         _ExtentY        =   6800
         _Version        =   393216
         Style           =   1
         Tabs            =   1
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
         TabCaption(0)   =   "&1. Item Detail"
         TabPicture(0)   =   "Invtrn.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fmeMsfdetail"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame fmeMsfdetail 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   12015
            Begin VB.PictureBox picInputRow 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   795
               Left            =   120
               ScaleHeight     =   765
               ScaleWidth      =   11745
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   120
               Width           =   11775
               Begin CommCtrls.CtxtBox txtUnit 
                  Height          =   375
                  Left            =   6000
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   1155
                  _ExtentX        =   2037
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
               Begin HlpN.HlpNCode hlpItem 
                  Height          =   375
                  Left            =   525
                  TabIndex        =   7
                  Top             =   360
                  Width           =   3990
                  _ExtentX        =   7038
                  _ExtentY        =   661
                  NameWidth       =   2600
               End
               Begin CommCtrls.ItxtBox txtSrno 
                  Height          =   375
                  Left            =   30
                  TabIndex        =   6
                  Top             =   375
                  Width           =   495
                  _ExtentX        =   873
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
               End
               Begin CommCtrls.ItxtBox txtQty 
                  Height          =   375
                  Left            =   4515
                  TabIndex        =   8
                  Top             =   375
                  Width           =   1455
                  _ExtentX        =   2566
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
               End
               Begin CommCtrls.NTxtBox txtRtl_Prc 
                  Height          =   375
                  Left            =   7155
                  TabIndex        =   10
                  Top             =   375
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  Desimal         =   3
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
               Begin CommCtrls.NTxtBox txtAmt 
                  Height          =   375
                  Left            =   8490
                  TabIndex        =   11
                  Top             =   375
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  Locked          =   -1  'True
                  Desimal         =   3
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
               Begin VB.CommandButton cmdOk 
                  Caption         =   "Ok"
                  Height          =   255
                  Left            =   9600
                  TabIndex        =   12
                  Top             =   480
                  Width           =   375
               End
               Begin VB.Label lblUnit 
                  Alignment       =   2  'Center
                  BackColor       =   &H00DCFBFC&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Unit"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   6000
                  TabIndex        =   20
                  Top             =   0
                  Width           =   1140
               End
               Begin VB.Label lblAmt 
                  Alignment       =   2  'Center
                  BackColor       =   &H00DCFBFC&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Amount"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   8490
                  TabIndex        =   22
                  Top             =   0
                  Width           =   1575
               End
               Begin VB.Label lblRatePerUnit 
                  Alignment       =   2  'Center
                  BackColor       =   &H00DCFBFC&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Rate/Unit"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   7155
                  TabIndex        =   21
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.Label lblPoPcs 
                  Alignment       =   2  'Center
                  BackColor       =   &H00DCFBFC&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Qty"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   4515
                  TabIndex        =   19
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.Label lblRawItem 
                  Alignment       =   2  'Center
                  BackColor       =   &H00DCFBFC&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Item Name"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   525
                  TabIndex        =   18
                  Top             =   0
                  Width           =   3990
               End
               Begin VB.Label lblSr 
                  Alignment       =   2  'Center
                  BackColor       =   &H00DCFBFC&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Sr"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   30
                  TabIndex        =   17
                  Top             =   0
                  Width           =   495
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
               Height          =   2355
               Left            =   120
               TabIndex        =   23
               Top             =   930
               Width           =   11775
               _ExtentX        =   20770
               _ExtentY        =   4154
               _Version        =   393216
               Rows            =   1
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   16308668
               ForeColorSel    =   0
               FocusRect       =   0
               HighLight       =   2
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   1
            End
            Begin VB.Label lblEccNo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4920
               TabIndex        =   36
               Top             =   840
               Width           =   975
            End
         End
      End
   End
   Begin HlpN.HlpNCode hlpFind 
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmInvtrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mEntryMode As String
Public mActCtrl As Control

Dim i As Integer

Const dColSrNo = 0
Const dColItm_id = 1
Const dColItm_name = 2
Const dColQty = 3
Const dColUnit = 4
Const dColrtl_rpc = 5
Const dColAmt = 6

Public Sub EntryAdd()
On Error GoTo errhndl
MP vbHourglass
    
    mEntryMode = "add"
    ClearScreen
    ClearMsf msfDetail
    SSTabDetail.Tab = 0

    
    EnableDisable True
    SetMsfDetail msfDetail
    txtVno.Text = GetMaxVno("Invtrn")
    SetFocusTo cmbType
        
    mskRec_Dat.Text = Date
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
    
End Sub

Public Sub EntryEdit(iViewMode As ViewMode)
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtVno.Text) <= 0 Then
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
    
    SetFocusTo txtDoc_No

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Public Sub EntryDelete()
On Error GoTo errhndl
MP vbHourglass
    
    If Val(txtVno.Text) <= 0 Then
        MsgBox "No Record Selected For Delete ", vbCritical
        BtnPressed mdiMainMenu.TbrMain.Buttons(btnCancel)
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    If MsgBox("Want to Delete EntryNo " & txtVno.Text & "..???", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        SetFocusTo SSTab1
        Exit Sub
    End If
    
    mEntryMode = "delete"
    
    SQL = "Delete from Invtrn"
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Vno= " & Val(txtVno.Text)
    gCnnMst.Execute SQL
    
    SQL = "Delete from Invdet"
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Vno= " & Val(txtVno.Text)
    gCnnMst.Execute SQL
    
    MsgBox "Entry No : " & txtVno.Text & " Deleted ", vbInformation
        
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
    ClearMsf msfDetail
    
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
        .TableName = GetDbTable("Invtrn", gMdbMst)
        .FieldList = "Vno,doc_no,Convert(Varchar(10),rec_dat,103) as ReceiveDate"
        .CodeField = "Vno"
        .NameField = "doc_no"
        .SetFocus
        .ShowHelp
    End With
    
    txtVno.Text = Val(hlpFind.CodeText)
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

Private Sub cmdOk_GotFocus()
    If LCase(lblMode.Caption) = "add" Or LCase(lblMode.Caption) = "edit" Then
        AddMsfdetail msfDetail
        If MsgBox("Want to add another Item...???", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            SetFocusTo txtSrno
        Else
            'SSTabDetail.Tab = 1
            SetFocusTo txtRemarks
        End If
    End If
End Sub

Private Sub Form_Activate()
On Error GoTo errhndl
    
    MP vbHourglass
    
    SetTextBoxes
    
    MP vbDefault
    
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub Form_Load()

On Error GoTo errhndl
MP vbHourglass
    
    GrabActiveControl
    SetMsfDetail msfDetail

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

Private Sub SetMsfDetail(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass
    
    With s_Msf
        .Cols = 7
        .RowHeightMin = 300
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(dColSrNo) = 560
        .ColAlignment(dColSrNo) = flexAlignLeftCenter
        
        .ColWidth(dColItm_id) = 1380
        .ColAlignment(dColItm_id) = flexAlignLeftCenter
        
        .ColWidth(dColItm_name) = 2600
        .ColAlignment(dColItm_name) = flexAlignLeftCenter
        
        .ColWidth(dColQty) = 1460
        .ColAlignment(dColQty) = flexAlignRightCenter
        
        .ColWidth(dColUnit) = 1200
        .ColAlignment(dColUnit) = flexAlignRightCenter
        
        .ColWidth(dColrtl_rpc) = 1330
        .ColAlignment(dColrtl_rpc) = flexAlignRightCenter
        
        .ColWidth(dColAmt) = 1550
        .ColAlignment(dColAmt) = flexAlignRightCenter
        
    End With

MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub SetTextBoxes()
        
    With hlpItem
        .SetAdoConnStr = gCnnMst
        .TableName = "Items"
        .FieldList = "code,name,shortname"
        .CodeField = "code"
        .NameField = "shortname"
        .DefaultSearchCol = 1
        .SetFontParameters "", gGujaratiFontName, gGujaratiFontName, 2, 12
    End With
        
    hlpFind.Visible = False
    
    FillTypeCombo
    
    VisibleNoVisibleBtn True
    SetActiveModeNControl mEntryMode
    CenterFrmChild Me
End Sub

Private Sub SaveInTmp()
On Error GoTo errhndl
Dim SQL As String
Dim i As Integer
MP vbHourglass
        
    gCnnMst.BeginTrans
    
    If LCase(lblMode.Caption) = "edit" Then
        SQL = "Delete from Invtrn "
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Vno = " & Val(txtVno.Text)
        gCnnMst.Execute SQL
        
        SQL = "Delete from Invdet "
        SQL = SQL & " Where 1=1"
        SQL = SQL & " And Vno = " & Val(txtVno.Text)
        gCnnMst.Execute SQL
    End If
    
    '---Insert into Invtrn
    SQL = "Insert into  Invtrn ("
    SQL = SQL & " Vno"
    SQL = SQL & ",ter_id"
    SQL = SQL & ",export_fg"
    
    SQL = SQL & ",tran_type"
    SQL = SQL & ",doc_no"
    SQL = SQL & ",rec_dat"
    SQL = SQL & ",remarks"
    
    SQL = SQL & ",dtadat "
    SQL = SQL & ",dtatim "
    SQL = SQL & ",dtausr "
    
    SQL = SQL & ",Trng_fg"
    
    SQL = SQL & " ) Values ("
    
    SQL = SQL & Val(txtVno.Text)
    SQL = SQL & "," & Val(gTerminalId)
    SQL = SQL & "," & "0"
    
    If cmbType.ListIndex <> -1 Then
        SQL = SQL & "," & Val(cmbType.ItemData(cmbType.ListIndex))
    Else
        SQL = SQL & "," & "2"   'Stock Inward
    End If
    
    SQL = SQL & "," & AQ(txtDoc_No.Text)
    SQL = SQL & "," & IIf(IsDate(mskRec_Dat.Text), ConvDatSql(mskRec_Dat.Text), "NULL")
    SQL = SQL & "," & AQ(txtRemarks.Text)
    
    SQL = SQL & "," & ConvDatSql(Date, BE_SQLSrv)
    SQL = SQL & "," & AQ(DtaTime)
    
    SQL = SQL & "," & AQ(gUser)
    
    SQL = SQL & "," & IsTrainingMode
    
    SQL = SQL & ")"
    gCnnMst.Execute SQL
    
    '---Entry in Invdet
    With msfDetail
        For i = 0 To .Rows - 1
            SQL = "Insert into  Invdet ("
            SQL = SQL & " vno "
            SQL = SQL & ",srno "
            SQL = SQL & ",itm_code "
            SQL = SQL & ",rtl_prc "
            
            SQL = SQL & ",qty "
            SQL = SQL & ",Amt "
            
            SQL = SQL & ",Trng_fg"
            
            SQL = SQL & " ) Values ("
            
            SQL = SQL & Val(txtVno.Text)
            SQL = SQL & "," & Val(.TextMatrix(i, dColSrNo))
            SQL = SQL & "," & Val(.TextMatrix(i, dColItm_id))
            SQL = SQL & "," & Val(.TextMatrix(i, dColrtl_rpc))
            
            Select Case cmbType.ItemData(cmbType.ListIndex)
                Case 1, 2, 3    ' +
                    SQL = SQL & "," & Val(.TextMatrix(i, dColQty))
                    SQL = SQL & "," & Val(.TextMatrix(i, dColAmt))

                Case 11, 12, 13    ' -
                    SQL = SQL & "," & Val(.TextMatrix(i, dColQty)) * -1
                    SQL = SQL & "," & Val(.TextMatrix(i, dColAmt)) * -1
            End Select
            
            SQL = SQL & "," & IsTrainingMode
    
            SQL = SQL & ")"
            gCnnMst.Execute SQL
        
        Next
    End With
    
    gCnnMst.CommitTrans

    MsgBox "Entry No : " & txtVno.Text & " Saved ", vbInformation
    
MP vbDefault
Exit Sub
errhndl:
    gCnnMst.RollbackTrans
    ErrMsg
    
End Sub

Private Sub Nevigate(s_Mode As Nevigate)
On Error GoTo errhndl
MP vbHourglass
    
    Dim rsttmp As ADODB.Recordset
        
    Select Case s_Mode
        Case MoveFirst
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " order by vno"
        Case MoveNext
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " Where 1 = 1 "
            SQL = SQL & " and vno > " & Val(txtVno.Text)
            SQL = SQL & " order by vno"
        Case MovePrev
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " Where 1 = 1 "
            SQL = SQL & " and vno < " & Val(txtVno.Text)
            SQL = SQL & " order by vno desc"
        Case MoveLast
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " order by vno desc"
        Case MoveTo
            SQL = "select top 1 * from " & GetDbTable("Invtrn", gMdbMst)
            SQL = SQL & " Where 1 = 1 "
            SQL = SQL & " and vno = " & Val(txtVno.Text)
            
    End Select
    
    OpenAdoRst rsttmp, SQL
        
        AdoRsRead rsttmp
        
        If rsttmp.RecordCount > 0 Then
            Select Case Val(rsttmp.Fields("tran_type"))
                Case 2
                    cmbType.ListIndex = 0
                Case 1
                    cmbType.ListIndex = 1
                Case 3
                    cmbType.ListIndex = 2
                Case 11
                    cmbType.ListIndex = 3
                Case 12
                    cmbType.ListIndex = 4
                Case 13
                    cmbType.ListIndex = 5
                Case Else
                    'do nothing
            End Select
        End If
        
        ReadInvdet rsttmp
        
        CloseAdoRst rsttmp
    
    CalculateTotal msfDetail
    
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next
End Sub

Private Sub ReadInvdet(s_rsttmp As ADODB.Recordset)
On Error GoTo errhndl
MP vbHourglass
    
    If s_rsttmp.RecordCount <= 0 Then
        MsgBox "No Records Found..", vbInformation
        SetFocusTo SSTab1
    Else
        fmeMsfdetail.Enabled = True
        
        Dim rsttmp As ADODB.Recordset
            
        SQL = "Select Invdet.Srno" & vbCrLf
        SQL = SQL & ",Invdet.Itm_code" & vbCrLf
        SQL = SQL & ",Items.ShortName " & vbCrLf
        SQL = SQL & ",Abs(Invdet.qty),units.name" & vbCrLf
        SQL = SQL & ",Invdet.rtl_prc,Abs(Invdet.amt)" & vbCrLf
        
        SQL = SQL & " From " & GetDbTable("Invdet", gMdbMst) & " Invdet " & vbCrLf
        SQL = SQL & " Left join " & GetDbTable("Items", gMdbMst) & " Items "
        SQL = SQL & " on (Invdet.itm_code=Items.code)" & vbCrLf
        
        SQL = SQL & " Left join " & GetDbTable("units", gMdbMst) & " units "
        SQL = SQL & " ON (Items.unit_id=units.code)" & vbCrLf
        
        SQL = SQL & " where Invdet.vno = " & Val(txtVno.Text) & vbCrLf
         
        OpenAdoRst rsttmp, SQL
        If rsttmp.RecordCount > 0 Then
            Set msfDetail.Recordset = rsttmp
        Else
            MsgBox "No Records Found For Detail Part", vbExclamation
            SetFocusTo SSTab1
        End If
        
        CloseAdoRst rsttmp
     End If
     
MP vbDefault
Exit Sub
errhndl:
    ErrMsg
    Resume Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, Shift, False
End Sub

Public Sub EnableDisable(s_Enable As Boolean)
    FmeCompanyDetail.Enabled = s_Enable
    fmeRecDetail.Enabled = s_Enable
    fmeTotals.Enabled = s_Enable
    fmeMsfdetail.Enabled = s_Enable
End Sub

Private Sub AddMsfdetail(s_Msf As MSHFlexGrid)
    
    With s_Msf
    
        If Val(txtSrno.Text) > .Rows Then
            .AddItem ""
        End If
        i = Val(txtSrno.Text) - 1
        
        .TextMatrix(i, dColSrNo) = Val(txtSrno.Text)
        .TextMatrix(i, dColItm_id) = Val(hlpItem.CodeText)
        .TextMatrix(i, dColItm_name) = Trim(hlpItem.NameText)
        .TextMatrix(i, dColQty) = Val(txtQty.Text)
        .TextMatrix(i, dColUnit) = txtUnit.Text
        .TextMatrix(i, dColrtl_rpc) = Val(txtRtl_Prc.Text)
        .TextMatrix(i, dColAmt) = Val(txtAmt.Text)
            
        InOrd_SrNo msfDetail
                    
        .Refresh
        .Row = .Rows - 1
        .TopRow = .Row
        .ColSel = .Cols - 1
        
        SetGridColGujFont s_Msf, dColItm_name, 12
        
        CalculateTotal msfDetail
        ClearInputLine
    End With
    
End Sub

Private Sub hlpItem_Validate(Cancel As Boolean)

    If Val(hlpItem.CodeText) <= 0 Then
        Cancel = True
    Else
        Dim rsttmp As ADODB.Recordset
        SQL = "select UM.code,UM.[name] "
        SQL = SQL & " from " & GetDbTable("Items", gMdbMst) & " AS RM"
        SQL = SQL & " inner join" & GetDbTable("Units", gMdbMst) & " AS UM"
        SQL = SQL & " ON UM.code=RM.code"
        SQL = SQL & " where RM.code=" & Val(hlpItem.CodeText)
        
        OpenAdoRst rsttmp, SQL, , , , gCnnMst
        If rsttmp.RecordCount > 0 Then
            txtUnit.Text = rsttmp.Fields("name")
            txtUnit.Tag = rsttmp.Fields("code")
        End If
    End If
End Sub

Private Sub msfDetail_Click()
On Error GoTo errhndl
 
MP vbHourglass
    
    With msfDetail
        If .Rows > 0 Then
            msfdetailToObj .Row
        Else
            If picInputRow.Enabled = True Then SetFocusTo txtSrno
        End If
    End With

MP vbDefault
 
Exit Sub
errhndl:
    Resume Next
End Sub

Private Sub msfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errhndl
 
MP vbHourglass
   
    msfDetail_Click
    
    If LCase(lblMode.Caption) = "add" Or LCase(lblMode.Caption) = "edit" Then
        Select Case KeyCode
            Case vbKeyDelete
                If MsgBox("Are you Sure you Want To Delete ?", vbYesNo) = vbYes Then
                    DeleteMsfRow msfDetail
                End If
                
            Case vbKeyReturn
                SetFocusTo hlpItem
            
            Case vbKeyInsert
                txtSrno.Text = msfDetail.Rows + 1
                SetFocusTo txtSrno
        
        End Select
    End If

MP vbDefault
Exit Sub
errhndl:
    Resume Next

End Sub

Private Sub txtRemarks_LostFocus()
    AskSave txtVno.Text, msfDetail, mEntryMode
End Sub

Private Sub txtSrno_GotFocus()
On Error GoTo errhndl

MP vbHourglass
    
    With msfDetail
        If LCase(lblMode.Caption) = "change" Then
            txtSrno.Text = Val(txtSrno.Text) + 1
        Else
            txtSrno.Text = .Rows + 1
        End If
    End With

MP vbDefault
 
Exit Sub
errhndl:
    Resume Next

End Sub

Private Sub txtSrno_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errhndl
 
MP vbHourglass
    
    With msfDetail
        If .Rows <= 0 Then Exit Sub
        If KeyCode = vbKeyF3 Then
            SetFocusTo msfDetail
            .Row = 0
            .ColSel = .Cols - 1
        End If
    End With

MP vbDefault
 
Exit Sub
errhndl:
    Resume Next
End Sub

Private Sub txtSrno_LostFocus()
On Error GoTo errhndl
MP vbHourglass
    
    With msfDetail
        If Val(txtSrno.Text) = 0 Or Val(txtSrno.Text) > .Rows Then
            txtSrno.Text = .Rows + 1
        ElseIf Val(txtSrno.Text) > 0 And Val(txtSrno.Text) <= .Rows Then
            .Row = txtSrno.Text - 1
            .TopRow = .Row
            .ColSel = .Cols - 1
            msfdetailToObj .Row
        End If
    End With

MP vbDefault
Exit Sub

errhndl:
    Resume Next
End Sub

Private Sub msfdetailToObj(s_Row As Integer)

    With msfDetail
        i = s_Row
        txtSrno.Text = .TextMatrix(i, dColSrNo)
        hlpItem.CodeText = .TextMatrix(i, dColItm_id)
        hlpItem.NameText = .TextMatrix(i, dColItm_name)
        txtQty.Text = .TextMatrix(i, dColQty)
        txtUnit.Text = .TextMatrix(i, dColUnit)
        txtRtl_Prc.Text = .TextMatrix(i, dColrtl_rpc)
        txtAmt.Text = .TextMatrix(i, dColAmt)
    End With
    
End Sub

Private Sub InOrd_SrNo(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass

    With s_Msf
        For i = 0 To .Rows - 1
            .Row = i
            .TextMatrix(i, dColSrNo) = i + 1
            .Refresh
        Next
    End With

MP vbDefault
Exit Sub
errhndl:
    Resume Next

End Sub

Private Sub DeleteMsfRow(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass

    With s_Msf
        If .Rows > 1 Then
            .RemoveItem .Row
            .Refresh
            InOrd_SrNo s_Msf
        Else
            .Rows = 0
            SetMsfDetail s_Msf
        End If
        
        .TopRow = .Rows - 1
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1

        SetFocusTo s_Msf
        
        CalculateTotal s_Msf
        
    End With

MP vbDefault

Exit Sub
errhndl:
    Resume Next
End Sub

Private Sub CalculateTotal(s_Msf As MSHFlexGrid)
On Error GoTo errhndl
MP vbHourglass
    
Dim mTotAmt As Double
Dim mRecPcs As Double

    With s_Msf
        For i = 0 To .Rows - 1
            mTotAmt = mTotAmt + Val(.TextMatrix(i, dColAmt))
            mRecPcs = mRecPcs + Val(.TextMatrix(i, dColQty))
        Next
    End With
    
    txtItemTot.Text = mTotAmt
    txtTotRecQty.Text = mRecPcs
    
MP vbDefault
Exit Sub
errhndl:
    Resume Next
    
End Sub

Private Sub CalculateWtItemTot()
On Error GoTo errhndl
MP vbHourglass
        
    txtAmt.Text = Val(txtQty.Text) * Val(txtRtl_Prc.Text)
    
MP vbDefault
Exit Sub
errhndl:
    Resume Next
End Sub

Private Sub txtRtl_Prc_LostFocus()
    CalculateWtItemTot
End Sub

Private Sub ClearMsf(s_Msf As MSHFlexGrid)
    With s_Msf
        .Clear
        .Rows = 0
        .Cols = 9
    End With
End Sub

Private Sub ClearInputLine()
    hlpItem.CodeText = 0
    hlpItem.NameText = ""
    txtQty.Text = 0
    txtUnit.Text = ""
    txtRtl_Prc.Text = 0
    txtAmt.Text = 0
End Sub


Private Sub FillTypeCombo()
    
    '   1   -   Opening Balance
    '   2   -   Stock Inward
    '   3   -   Stock Adjustment Up
    '   4   -   Receive from Store
    
    '   11  -   Stock Adjustment Down
    '   12  -   Stock Waste
    '   13  -   Issue For Sale
    
    
    With cmbType
        .AddItem "Opening Stock - 1"
        .ItemData(.NewIndex) = 1
        
        .AddItem "Stock Inward - 2"
        .ItemData(.NewIndex) = 2
        
        .AddItem "Stock Adjustment Up - 3"
        .ItemData(.NewIndex) = 3
    
        .AddItem "Receive from Store - 4"
        .ItemData(.NewIndex) = 4
    
        .AddItem "Stock Adjustment Down - 11"
        .ItemData(.NewIndex) = 11
        
        cmbType.AddItem "Stock Waste - 12"
        .ItemData(.NewIndex) = 12
        
        cmbType.AddItem "Issue for Sale - 13"
        .ItemData(.NewIndex) = 13
        
        .ListIndex = 1
    End With
End Sub

Private Sub SSTabDetail_Click(PreviousTab As Integer)
    Select Case Val(SSTabDetail.Tab)
        Case 0
            If fmeMsfdetail.Enabled Then SetFocusTo msfDetail
        Case 1
            'do nothing
    End Select
End Sub

