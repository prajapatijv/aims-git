VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   600
   End
   Begin VB.Shape ShapeBox1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      FillColor       =   &H00F8D9BC&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   3840
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape shapeBox 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   240
      Top             =   2160
      Width           =   135
   End
   Begin VB.Line LineBoxBottom 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   1080
      X2              =   3960
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line LineBoxTop 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   3480
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AIMS : Inventory Management System"
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
      Left            =   780
      TabIndex        =   9
      Top             =   120
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   6240
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   5520
      X2              =   5520
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AIMS : Inventory Management System"
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
      Left            =   810
      TabIndex        =   8
      Top             =   120
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2490
      Left            =   15
      Top             =   15
      Width           =   6480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Released Version:1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   1740
      TabIndex        =   7
      Top             =   1680
      Width           =   2265
   End
   Begin VB.Label lblAni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblAni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblAni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblAni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblAni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblAni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Software"
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
      Height          =   495
      Left            =   1710
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00DCFBFC&
      Height          =   1935
      Left            =   50
      TabIndex        =   10
      Top             =   520
      Width           =   5400
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnt As Integer
Dim i As Integer
Dim mClr As OLE_COLOR
Dim mClrBlink As OLE_COLOR

Private Sub Form_Activate()

    TransperentForm Me, 100
    
    mClr = &H80000005  ' &HDCFBFC   '&H80&
    mClrBlink = &HF8D9BC
    lblAni(0).BackColor = mClrBlink
    
End Sub

Private Sub Form_Click()
   mdiMainMenu.Show
   Unload frmSplash
End Sub

Private Sub Form_Load()
    Timer1.Interval = 500
    Timer1_Timer
    lblMsg.Caption = "Loading Software" & vbCrLf & "Components"
End Sub

Private Sub Label5_Click()
   mdiMainMenu.Show
   Unload frmSplash
End Sub

Private Sub Timer1_Timer()
     DoEvents
     If lblAni(i).BackColor = mClr Then
        lblAni(0).BackColor = mClr
        lblAni(1).BackColor = mClr
        lblAni(2).BackColor = mClr
        lblAni(3).BackColor = mClr
        lblAni(4).BackColor = mClr
        lblAni(5).BackColor = mClr
        lblAni(i).BackColor = mClrBlink  '&H40C0&
     Else
        lblAni(i).BackColor = mClrBlink
     End If
     
     shapeBox.Left = shapeBox.Left + 200
     ShapeBox1.Left = ShapeBox1.Left - 200
    
     i = i + 1
     
     TransperentForm Me, 100 + (i * 25)
     
     If i > 5 Then
        cnt = cnt + 1
        If cnt >= 3 Then Form_Click
        i = 0
     End If
    
     
End Sub
