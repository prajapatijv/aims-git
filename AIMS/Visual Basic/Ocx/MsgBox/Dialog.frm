VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   0  'None
   Caption         =   "MsgBox"
   ClientHeight    =   3405
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   7515
   Icon            =   "Dialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Dialog.frx":0E42
   ScaleHeight     =   3405
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Blink 
      Left            =   120
      Top             =   2880
   End
   Begin VB.CommandButton Button3 
      BackColor       =   &H008080FF&
      Caption         =   "Retry"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1215
   End
   Begin VB.CommandButton Button2 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1215
   End
   Begin VB.CommandButton Button1 
      BackColor       =   &H008080FF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Label lblHeadF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   460
      TabIndex        =   6
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label LblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6900
      TabIndex        =   4
      Top             =   10
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Dialog.frx":4A2C
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   960
      TabIndex        =   3
      Top             =   840
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblHeadB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   420
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mPos As Integer
Dim mLineNo As Integer

Public Sub BtnClick()
    gBtnClicked = "jv" & LCase(Screen.ActiveControl.Caption)
    Unload Me
End Sub

Private Sub Blink_Timer()
    lblMsg.Visible = Not lblMsg.Visible
End Sub

Private Sub Button1_Click()
    BtnClick
End Sub

Private Sub Button2_Click()
    BtnClick
End Sub

Private Sub Button3_Click()
    BtnClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyPageDown
            mPos = InStr(mPos + 1, gMsgStr, vbCrLf, vbTextCompare)
            If mPos > 0 Then
                mLineNo = mLineNo + 1
                lblMsg.Caption = "...(" & mLineNo + 1 & ")... " & Trim(Mid(gMsgStr, mPos + 2, Len(gMsgStr)))
            Else
                mLineNo = 0
                mPos = 1
                lblMsg.Caption = Trim(Mid(gMsgStr, mPos, Len(gMsgStr)))
            End If

        'Case vbKeyPageUp
        
        Case vbKeyEscape
            LblExit_Click
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    Blink_Timer
    
'    MsgBox mDefaBtn
'    Select Case Val(mDefaBtn)
'        Case 0
'            If Button1.Visible = True Then Button1.TabIndex = 0
'        Case 1
'            If Button2.Visible = True Then Button2.TabIndex = 0
'        Case 2
'            If Button3.Visible = True Then Button3.TabIndex = 0
'    End Select
End Sub

Private Sub LblExit_Click()
    Unload Me
End Sub

