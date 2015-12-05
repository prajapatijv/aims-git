VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Begin VB.Form frmQty 
   BorderStyle     =   0  'None
   Caption         =   "Quantity"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   Icon            =   "qty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
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
         Left            =   2400
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1020
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
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin CommCtrls.NTxtBox txtPayAmt 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
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

Dim dMindTicketAmout As Double
Dim dTicketAmout As Double
Dim btnOkPressed As Boolean

Public Function Display(s_dTicketAmount As Double, ByRef s_btnOkPressed As Boolean) As Double
    
    txtPayAmt.Text = s_dTicketAmount
    dMindTicketAmout = s_dTicketAmount
    
    With frmPosGui
        Me.Move .frmMainContainer.Left + (.mshTicket.Width + 250), _
                            (.frmMainContainer.Top + 300)
    End With
    
    Me.Show vbModal
    s_btnOkPressed = btnOkPressed
    Display = dTicketAmout
    
End Function

Private Sub SetTextBoxes()
    
    txtPayAmt.Font.Name = gGujaratiFontName
    txtPayAmt.Font.Size = 18
    txtPayAmt.Font.Bold = True
    
End Sub

Private Sub cmdCancel_Click()
    dTicketAmout = 0
    btnOkPressed = False
    Unload Me
End Sub

Private Sub cmdDenoms_Click(Index As Integer)
    
    cmdOk_Click
    
End Sub

Private Sub cmdOk_Click()

    If Val(txtPayAmt.Text) < dMindTicketAmout Then
        MsgBox "Minimum Ticket Amount " & dMindTicketAmout
        txtPayAmt.Text = dMindTicketAmout
        Exit Sub
    End If
    
    dTicketAmout = Val(txtPayAmt.Text)
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
