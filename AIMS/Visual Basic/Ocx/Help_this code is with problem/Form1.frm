VERSION 5.00
Object = "{F1857142-34CB-11D9-91C4-A0CC4AC10000}#20.0#0"; "MyHelp.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin HlpN.HlpNCode hlpConum 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   661
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const gCnninv = "dsn_mst" '"Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=HMst;Data Source=g2;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=G2;Use Encryption for Data=False;Tag with column collation when possible=False"

Private Sub Command1_Click()
    hlpConum.GetNameText hlpConum.CodeText
End Sub

Private Sub Form_Activate()
    With hlpConum
        .SetAdoConnStr = gCnninv
        .TableName = "CompMast"
        .FieldList = "code,name"
        .CodeField = "code"
        .NameField = "Name"
        .SqlWhere = " Status = 'T'"
    End With
End Sub

Private Sub Form_DblClick()
    HlpNCode1.GetNameText (HlpNCode1.CodeText)
End Sub

