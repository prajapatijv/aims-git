VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Begin VB.Form frmInvtrnDetails 
   Caption         =   "Inventory Entry"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5940
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInputRow 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   120
      ScaleHeight     =   2805
      ScaleWidth      =   5625
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   5655
      Begin CommCtrls.NTxtBox txtQty 
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Desimal         =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin CommCtrls.NTxtBox txtRtl_Prc 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Desimal         =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowNull       =   -1  'True
      End
      Begin CommCtrls.NTxtBox txtAmt 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Locked          =   -1  'True
         BackColor       =   14482428
         Desimal         =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
         AllowNull       =   -1  'True
      End
      Begin CommCtrls.ItxtBox txtItemName 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         Locked          =   -1  'True
         BackColor       =   14482428
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CommCtrls.ItxtBox txtItemCode 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.Label lblRawItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblPoPcs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   90
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblRatePerUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   90
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblAmt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   90
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmInvtrnDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const dColItm_id = 0
Const dColItm_name = 1
Const dColQty = 2
Const dColUnit = 3
Const dColrtl_rpc = 4
Const dColAmt = 5
Const dColTranType = 6

Dim btnOkPressed As Boolean
Dim arr(6) As String
Dim tranType As Integer
Dim zerobarredTypes() As String

Private Sub CalculateWtItemTot()
On Error GoTo errhndl
MP vbHourglass
        
    txtAmt.Text = Val(txtQty.Text) * Val(txtRtl_Prc.Text)
    
MP vbDefault
Exit Sub
errhndl:
    Resume Next
End Sub


Private Sub SetFieldValues(s_arrdata() As String)

    Dim iCnt As Integer
    
    For iCnt = 0 To UBound(s_arrdata)
        arr(iCnt) = s_arrdata(iCnt)
    Next
    
    txtItemCode.Text = arr(dColItm_id)
    txtItemName.Text = arr(dColItm_name)
    txtQty.Text = arr(dColQty)
    txtRtl_Prc.Text = arr(dColrtl_rpc)
    txtAmt.Text = arr(dColAmt)
    tranType = arr(dColTranType)
    
    If Not ZeroUnitPriceAllowed() Then
        txtRtl_Prc.AllowNull = False
    Else
        txtRtl_Prc.AllowNull = True
    End If
    
End Sub

Public Function Display(s_arrdata() As String)

    SetFieldValues s_arrdata()
    
    Me.Show vbModal
    
    Display = arr
    
End Function

Private Sub cmdCancel_Click()
    btnOkPressed = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    btnOkPressed = True
    
    arr(dColQty) = txtQty.Text
    arr(dColrtl_rpc) = txtRtl_Prc.Text
    arr(dColAmt) = txtAmt.Text

    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, Shift, False
End Sub

Private Sub Form_Load()
    
    txtItemName.Font.Name = gGujaratiFontName
    txtItemName.Font.Size = 12
    txtItemName.Locked = True
    
    zerobarredTypes = Split(gDenyZeroPriceMaterialInwardOutwardTypes, ",")
    
    CenterFrmNonChild Me
    
End Sub

Private Function ZeroUnitPriceAllowed() As Boolean
    
    Dim iCnt As Integer
    
    For iCnt = 0 To UBound(zerobarredTypes)
        If tranType = zerobarredTypes(iCnt) Then
            ZeroUnitPriceAllowed = False
            Exit Function
        End If
    Next
    
    ZeroUnitPriceAllowed = True
    Exit Function
End Function


Private Sub txtQty_Change()

    CalculateWtItemTot
    
End Sub


Private Sub txtRtl_Prc_Change()
    
    CalculateWtItemTot
    
End Sub

Private Sub txtRtl_Prc_LostFocus()
    
    CalculateWtItemTot
    
End Sub
