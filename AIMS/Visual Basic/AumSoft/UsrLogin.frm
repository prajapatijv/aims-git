VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Begin VB.Form frmUsrLogin 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   Caption         =   "User Login"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame frmLogin 
      BackColor       =   &H00DCFBFC&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   60
      TabIndex        =   8
      Top             =   570
      Width           =   4815
      Begin VB.TextBox txtPwd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin CommCtrls.CtxtBox txtUName 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Alignment       =   0
         MaxLength       =   15
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
      Begin VB.Label lblPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   900
      End
      Begin VB.Label lblUName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   2085
      Left            =   0
      Top             =   0
      Width           =   6330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
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
      Left            =   3570
      TabIndex        =   6
      Top             =   120
      Width           =   1530
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   5040
      X2              =   5040
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   0
      X2              =   6000
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "frmUsrLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOk_Click()
    Dim rst As ADODB.Recordset
    
    SQL = "Select Uid from UserMast"
    SQL = SQL & " Where UName = " & AQ(txtUName.Text)
    SQL = SQL & " And Pwd = " & AQ(ChartoAsc(txtPwd.Text))
    
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    If rst.RecordCount > 0 Then
        gUser = rst.Fields("uid").Value
        gUserName = txtUName.Text
        
        mdiMainMenu.StatBar.Panels(2).Text = IIf(OperaionMode = enServer, "SERVER", "TERMINAL")
        mdiMainMenu.StatBar.Panels(3).Text = gUser & " : " & gUserName
        mdiMainMenu.StatBar.Panels(4).Text = "Terminal : " & gTerminalId
        
        SetUserModeMenu
        
        LoadProject

        Unload Me

    Else
        MsgBox "Invalid UserName/Passward...!!!", vbCritical
    End If
    
End Sub

Private Sub Form_Activate()
    CenterFrmNonChild Me
    
'    txtUName.Text = "ADMIN"
'    txtPwd.Text = "ok"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FormKeyDown KeyCode, 0, False
End Sub

Private Sub txtPwd_GotFocus()
    txtPwd.SelStart = 0
    txtPwd.SelLength = Len(txtPwd.Text)
    cmdOk.Default = True
End Sub

Private Sub txtUName_GotFocus()
    cmdOk.Default = False
End Sub

Private Sub LoadProject()
    
    CreatePkgTrnTables
    AddPkgConstraintsTrn
    
    If InStr(1, Command$, "install=check", vbTextCompare) > 0 Then
        chkStructure
    End If
    
    DeleteTrainingModeData
  
End Sub


Private Sub SetUserModeMenu()

    ReSetMenus

    Select Case UserLevel(gUser)
    
        Case enUserLevel.eAdmin                'ADMIN
            Select Case OperaionMode
                Case enOperationMode.enServer
                    With mdiMainMenu
                        .mnuTransArr1(1).Visible = True
                    End With
                Case enOperationMode.enTerminal
                    With mdiMainMenu
                        .mnuTransArr1(1).Visible = False
                    End With
            End Select
            
        Case enUserLevel.eImpex                 'IMPEX
            With mdiMainMenu
                .mnuMasters.Visible = False
                .mnuTrans.Visible = False
                
                .mnuAdminArr1(0).Enabled = False
            End With
            
        Case enUserLevel.eTraining              'Training
            With mdiMainMenu
                If OperaionMode = enServer Then
                    .mnuMasters.Visible = True
                Else
                    .mnuMasters.Visible = False
                End If
                
                .mnuUtility.Visible = False
                .mnuAdmin.Visible = False
                
            End With
            
        Case Else                               'Normal User
            Select Case OperaionMode
                Case enOperationMode.enServer
                    With mdiMainMenu
                        .mnuAdmin.Visible = False
                    End With
                
                Case enOperationMode.enTerminal
                    With mdiMainMenu
                        .mnuMasters.Visible = False
                        .mnuAdmin.Visible = False
                        
                        .mnuTransArr1(1).Enabled = False
                    End With
            End Select
            
    End Select
    
End Sub

''This Procedure enables all the menus and menu trees
Private Sub ReSetMenus()
    
    Dim i As Integer

    With mdiMainMenu
        ''Transction
        .mnuTrans.Visible = True
        For i = 0 To .mnuTransArr1.Count - 1
            .mnuTransArr1(i).Visible = True
        Next
    
        ''Master
        .mnuMasters.Visible = True
        For i = 0 To .mnuMasterArr1.Count - 1
            .mnuMasterArr1(i).Visible = True
        Next
    
        ''Reports
        .mnuReports.Visible = True
        For i = 0 To .mnuReportsArr1.Count - 1
            .mnuReportsArr1(i).Visible = True
        Next
    
        ''Admin
        .mnuAdmin.Visible = True
        For i = 0 To .mnuAdminArr1.Count - 1
            .mnuAdminArr1(i).Visible = True
        Next
        
        ''Utility
        .mnuUtility.Visible = True
        For i = 0 To .mnuUtilityArr1.Count - 1
            .mnuUtilityArr1(i).Visible = True
        Next
        
    End With
End Sub
