VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Begin VB.Form frmQty 
   BorderStyle     =   0  'None
   Caption         =   "Quantity"
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   Icon            =   "qty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6255
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Default         =   -1  'True
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
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   4920
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin CommCtrls.NTxtBox txtQty 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   873
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
   End
   Begin VB.Label lblItemName 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "frmQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mQty As Integer
Dim btnOkPressed As Boolean

Public Function Display(s_itemDescr As String, s_mInitQty As Integer, ByRef s_btnOkPressed As Boolean) As Integer
    
    lblItemName.Caption = s_itemDescr
    txtQty.Text = s_mInitQty
    
    With frmPosGui
        Me.Move .frmMainContainer.Left + (.mshTicket.Width + 250), _
                            (.frmMainContainer.Top + 300)
    End With
    
    Me.Show vbModal
    s_btnOkPressed = btnOkPressed
    Display = mQty
    
End Function

Private Sub SetTextBoxes()
    
    lblItemName.Font.Name = gGujaratiFontName
    lblItemName.Font.Size = 18
    lblItemName.Font.Bold = True
    
    txtQty.Font.Name = gGujaratiFontName
    txtQty.Font.Size = 18
    txtQty.Font.Bold = True
    
End Sub

Private Sub cmdCancel_Click()
    btnOkPressed = False
    Unload Me
End Sub

Private Sub cmdOk_Click()

    If Val(txtQty.Text) < 0 Then
        MsgBox "Qty must be positive value!"
        Exit Sub
    End If
    
    mQty = Val(txtQty.Text)
    btnOkPressed = True
    Unload Me

End Sub

Private Sub Form_Load()

    SetTextBoxes
    
End Sub

Private Sub Form_Resize()
    With Shape1
        .BorderWidth = 5
        .Move 0, 0, Me.Width, Me.Height
    End With
End Sub

