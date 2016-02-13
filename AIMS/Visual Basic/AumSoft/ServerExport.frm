VERSION 5.00
Begin VB.Form frmServerExport 
   BackColor       =   &H00F8D9BC&
   BorderStyle     =   0  'None
   Caption         =   "Export "
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   Icon            =   "ServerExport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
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
      Left            =   5880
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   5880
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame fmeReport 
      BackColor       =   &H00DCFBFC&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin VB.Frame fmeImport 
         BackColor       =   &H00DCFBFC&
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   5295
         Begin VB.Label lblImport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ">>>"
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
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fmeExport 
         BackColor       =   &H00DCFBFC&
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   5295
         Begin VB.CheckBox chkResetTerminalTables 
            Appearance      =   0  'Flat
            BackColor       =   &H00DCFBFC&
            Caption         =   "Truncate Termial Tables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblExport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ">>>"
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
            TabIndex        =   2
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   360
      Left            =   0
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Export Data"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   6840
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   5760
      X2              =   5760
      Y1              =   120
      Y2              =   3360
   End
   Begin VB.Label lblRepHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Export Data"
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
      Left            =   3270
      TabIndex        =   6
      Top             =   135
      Width           =   2640
   End
End
Attribute VB_Name = "frmServerExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mExportMdbPath As String
Dim mImportMdbPath As String
Dim DBname As String

Private Sub cmdExport_Click()

On Error GoTo Err:

    MP vbArrowHourglass
    
    cmdExport.Enabled = False
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    SQL = " Select * from ServerExport  "
    If OperaionMode = enServer Then
        SQL = SQL & " Where TableType = 'Server'"
    Else
        SQL = SQL & " Where TableType = 'Terminal'"
    End If
    SQL = SQL & " And actv_fg = 1"
    
    OpenAdoRst rst, SQL
    rst.MoveFirst
    
    While Not rst.EOF
    
        lblExport.Caption = " >>> Exporting Table " & vbCrLf & Space(9) & _
                                rst.Fields("TableName").Value & " . . ."
        
        ExpSql2Mdb rst.Fields("TableName").Value, rst.Fields("TableName").Value
        
        rst.MoveNext
                
        DoEvents
    Wend
    
    MsgBox "Export completed to " & vbCrLf & Trim$(mExportMdbPath)
    lblExport.Caption = "Export completed to " & vbCrLf & Trim$(mExportMdbPath)
    
    cmdExport.Enabled = True
    
    MP vbDefault
    
    rst.Close
    Set rst = Nothing
    Exit Sub
Err:
    MP vbDefault
    MsgBox Err.Number & vbCrLf & Err.Description
        
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdImport_Click()
    
    If OperaionMode = enTerminal Then
        ImportServerDataAtTerminal
    Else
        ImportTerminalDataAtServer
    End If
    
End Sub

Private Sub Form_Activate()
    
    SetTextBoxes
    
End Sub

Private Sub Form_Load()
    
    CenterFrmChild Me
    
End Sub

Private Sub Form_Resize()

    With Shape1
        .BorderWidth = 5
        .Move 0, 0, Me.Width, Me.Height
    End With

End Sub

Private Sub SetTextBoxes()

    mExportMdbPath = Readfromini("ExportMdbPath", App.Path + "\" + SetPkgName)
    mImportMdbPath = Readfromini("ImportMdbPath", App.Path + "\" + SetPkgName)
    
    CreateFolderHerarchy

    lblImport.Caption = "Import Data From " & vbCrLf & mImportMdbPath
    
    Select Case LCase$(Me.Tag)
        Case "export"
            cmdExport.Visible = True
            cmdImport.Visible = False
            cmdExport.Move cmdImport.Left, cmdImport.Top
            
            fmeExport.Visible = True
            fmeImport.Visible = False
            fmeExport.Move 120, 60, 5295, 2450
            
            chkResetTerminalTables.Visible = False
            'chkResetTerminalTables.Value = vbChecked
            'chkResetTerminalTables.Move 120, 2100
            
            If OperaionMode = enServer Then
                Label1.Caption = "Server Export Data"
                lblRepHead.Caption = Label1.Caption
            Else
                Label1.Caption = "Terminal Export Data"
                lblRepHead.Caption = Label1.Caption
            End If
            
        Case "import"
            cmdExport.Visible = False
            cmdImport.Visible = True
            fmeExport.Visible = False
            fmeImport.Visible = True
            fmeImport.Move 120, 60, 5295, 2450
            
            chkResetTerminalTables.Visible = False
            chkResetTerminalTables.Value = vbUnchecked
            
            If OperaionMode = enServer Then
                Label1.Caption = "Server Import Data"
                lblRepHead.Caption = Label1.Caption
            Else
                Label1.Caption = "Terminal Import Data"
                lblRepHead.Caption = Label1.Caption
            End If
            
    End Select
End Sub

Private Function ImportServerDataAtTerminal()

 On Error GoTo Err:
 
    MP vbArrowHourglass
    cmdImport.Enabled = False

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    SQL = " Select * from ServerExport  "
    SQL = SQL & " Where TableType = 'Server'"
    SQL = SQL & " And actv_fg = 1"
    
    OpenAdoRst rst, SQL
    rst.MoveFirst
    While Not rst.EOF
    
        lblImport.Caption = " >>> Importing Table " & vbCrLf & Space(9) & _
                                rst.Fields("TableName").Value & " . . ."
        
        ImportFromMdb2Sql rst.Fields("TableName").Value, "ServerImpex.mdb"
        
        rst.MoveNext
        
        DoEvents
    Wend

    'Move Mdb File to Done folder
    MoveMdbfile "ServerImpex.mdb"

    MsgBox "Import completed successfully."
    lblImport.Caption = "Import completed successfully."

    cmdImport.Enabled = True
    
    Set rst = Nothing
    
    Exit Function
Err:
    MP vbDefault
    MsgBox Err.Number & vbCrLf & Err.Description

End Function

Private Function ImportTerminalDataAtServer()

 On Error GoTo Err:
 
    MP vbArrowHourglass
    
    Dim mTerImportFileName As String
    
    While Dir$(mImportMdbPath & "TerminalImpex****.mdb") <> ""
    
        'Import Terminal Files from Terminal to Server
        Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    
        Dim rstTer As ADODB.Recordset
        Set rstTer = New ADODB.Recordset
    
        Dim importBarred As Boolean
        importBarred = False
        
        mTerImportFileName = Dir$(mImportMdbPath & "TerminalImpex****.mdb")
    
        SQL = " Select * from ServerExport  "
        SQL = SQL & " Where TableType = 'Terminal'"
        SQL = SQL & " And actv_fg = 1"

        OpenAdoRst rst, SQL
        rst.MoveFirst
        While Not rst.EOF

            SQL = "Select ImportBarred from TerminalConfig where code = " & Replace(Replace(mTerImportFileName, "TerminalImpex", ""), ".mdb", "")
            OpenAdoRst rstTer, SQL
            rstTer.MoveFirst
            importBarred = IIf(rstTer.Fields("ImportBarred").Value = True, True, False)

            If importBarred = False Then
                lblImport.Caption = " >>> Importing Table " & vbCrLf & Space(9) & _
                                        rst.Fields("TableName").Value & " . . ."
                
                ImportFromMdb2Sql rst.Fields("TableName").Value, mTerImportFileName
                
    
                rst.MoveNext
    
                DoEvents
            Else
                lblImport.Caption = " >>> Importing Barred for Table " & vbCrLf & Space(9) & _
                                        rst.Fields("TableName").Value & " . . ."
                                        
                DoEvents
            End If
        Wend

        'Move Mdb File to be Done
        MoveMdbfile mTerImportFileName
        
        rst.Close
        Set rst = Nothing
        ''''
    Wend

    MsgBox "Import completed successfully."
    lblImport.Caption = "Import completed successfully."

    cmdImport.Enabled = True
    
    Set rst = Nothing
    
    Exit Function
Err:
    MP vbDefault
    MsgBox Err.Number & vbCrLf & Err.Description

End Function


Private Sub CreateFolderHerarchy()

    On Error GoTo Err:
    
    MkDir mImportMdbPath & "Done"
    
Err:
    If Err.Number = 75 Then
        Resume Next
    End If
End Sub

Private Sub MoveMdbfile(s_FileName As String)

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    If fso.FileExists(mImportMdbPath & s_FileName) Then
        If fso.FileExists(mImportMdbPath & "Done\" & s_FileName) Then
            Call fso.DeleteFile(mImportMdbPath & "Done\" & s_FileName, True)
        End If
        fso.MoveFile mImportMdbPath & s_FileName, mImportMdbPath & "Done\" & s_FileName
    End If
    
    Set fso = Nothing
     
End Sub

