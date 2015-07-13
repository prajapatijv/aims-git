VERSION 5.00
Object = "{86144B5E-6628-49BD-BDDD-F6C4F692705D}#1.1#0"; "MyHelp.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin HlpN.HlpNCode HlpNCode3 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   4395
      _extentx        =   7752
      _extenty        =   661
   End
   Begin HlpN.HlpNCode HlpNCode2 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   4395
      _extentx        =   7752
      _extenty        =   661
   End
   Begin HlpN.HlpNCode HlpNCode1 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   4395
      _extentx        =   7752
      _extenty        =   661
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Const gCnninv = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AIMS_DB;Data Source=VGT-BHARAT\SQLEXPRESS"

Private Sub Command1_Click()
    HlpNCode1.ShowHelp
End Sub

Private Sub Command2_Click()
    MsgBox HlpNCode1.TextMatrixData
End Sub

Private Sub Form_Activate()
    With HlpNCode1
'        .SetAdoConnStr = gCnninv
        .TextMatrixDataCol = 2
        .SetAdoDSN = "AIMS_DSN"
        .TableName = "Items"
        .FieldList = "code,name,shortname"
        .CodeField = "code"
        .NameField = "ShortName"
        .SetFontParameters "", "Kanaiya-Normal", "Kanaiya-Normal", "2", 12
        .DefaultSearchCol = 2
    End With
    
    With HlpNCode2
        .SetAdoDSN = "AIMS_DSN"
        .TableName = "Sizes"
        .FieldList = "code,name"
        .CodeField = "code"
        .NameField = "name"
        .DefaultSearchCol = 1
    End With
    
    With HlpNCode3
        .SetAdoDSN = "AIMS_DSN"
        .TableName = "Sizes"
        .FieldList = "code,shortname,name"
        .CodeField = "code"
        .NameField = "shortname"
        .SetFontParameters "Kanaiya-Normal", "Kanaiya-Normal", "Kanaiya-Normal", "0~1", 12
        .DefaultSearchCol = 1
        
    End With
    
End Sub

