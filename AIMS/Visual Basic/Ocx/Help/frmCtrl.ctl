VERSION 5.00
Begin VB.UserControl HlpNCode 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   LockControls    =   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   735
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdValidate 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtName 
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
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2640
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "HlpNCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_CodeValidate = False
Const m_def_CodeField = 0
Const m_def_NameField = 0
Const m_def_SetAdoConnStr = 0
Const m_def_CodeGotFocusColor = 16047579
Const m_def_CodeLostFocusColor = 16777215
Const m_def_NameVisible = True
Const m_def_NameWidth = 3000
Const m_def_Topn = 100
Const m_def_CodeWidth = 1000
Const m_def_SpcBtwn = 20
Const m_def_GetCode = ""
Const m_def_GetName = ""
Const m_def_SqlSelect = ""
Const m_def_SqlWhere = ""
Const m_def_FieldList = ""
Const m_def_TableName = ""
Const m_def_SetAdoDSN = ""

'Property Variables:
Dim m_CodeValidate As Boolean
Dim m_CodeField As Variant
Dim m_NameField As Variant
Dim m_SetAdoConnStr As Variant
Dim m_SetAdoDSN As String

Dim m_CodeGotFocusColor As OLE_COLOR
Dim m_CodeLostFocusColor As OLE_COLOR
Dim m_NameVisible As Boolean
Dim m_NameWidth As Integer
Dim m_TopN As Integer
Dim m_CodeWidth As Integer
Dim m_SpcBtwn As Integer
Dim m_GetCode As String
Dim m_GetName As String
Dim m_SqlSelect As String
Dim m_SqlWhere As String
Dim m_FieldList As String
Dim m_TableName As String

'Font Settings Variable
Dim m_CodeFontName As String
Dim m_NameFontName As String
Dim m_GridFontName As String
Dim m_Gridcols As String
Dim m_FontSize As Integer

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
''Project variable

Private Sub cmdValidate_Click()
Dim SQL  As String
Dim rsttemp As New ADODB.Recordset
    
    txtCode_GotFocus
    txtCode_LostFocus
    
    If gAdoConnStr = "" Then
        gAdoConnStr = gSetAdoDSN
    End If
    
    If Val(txtCode.Text) <= 0 Then
        rsttemp.Open xSource, gAdoConnStr, adOpenStatic, adLockOptimistic
    Else
        SQL = " Select "
        If Val(m_TopN) > 0 Then
            SQL = SQL & " Top " & TopN
        End If
        SQL = SQL & " * "
        SQL = SQL & " from " & m_TableName
        SQL = SQL & " Where 1=1 "
        SQL = SQL & " And "
        SQL = SQL & m_CodeField & " = " & Val(txtCode.Text)
        
        If Len(m_SqlWhere) > 0 Then
            SQL = SQL & " And " & m_SqlWhere
        End If
        
        rsttemp.Open SQL, gAdoConnStr, adOpenDynamic, adLockOptimistic
        If Not rsttemp.BOF And Not rsttemp.EOF Then
            txtCode = rsttemp.Fields(f_CodeField)
            txtName = rsttemp.Fields(f_NameField)
            Exit Sub
        Else
            If rsttemp.State = adStateOpen Then rsttemp.Close
            rsttemp.Open xSource, gAdoConnStr, adOpenStatic, adLockOptimistic
        End If
    End If
    
    txtCode.Text = ""
    txtName.Text = ""
    
    frmHlp.msfHelp.Clear
    
    Set frmHlp.msfHelp.Recordset = rsttemp
    frmHlp.SetMsfDetail frmHlp.msfHelp
    
    If rsttemp.State = adStateOpen Then rsttemp.Close
    Set rsttemp = Nothing

    frmHlp.Show vbModal
    
End Sub

Private Sub txtCode_GotFocus()
    
    If Trim$(m_CodeFontName) = "" Then m_CodeFontName = "Arial"
    If Trim$(m_NameFontName) = "" Then m_NameFontName = "Arial"
    If Val(m_FontSize) = 0 Then m_FontSize = 10
    
    txtCode.Font.Name = m_CodeFontName
    txtName.Font.Name = m_NameFontName
    
    If LCase(txtCode.Font.Name) = LCase("Arial") Then
        txtCode.Font.Size = 10
    Else
        txtCode.Font.Size = m_FontSize
    End If
    If LCase(txtName.Font.Name) = LCase("Arial") Then
        txtName.Font.Size = 10
    Else
        txtName.Font.Size = m_FontSize
    End If
    
    'Reset Global Variables to current instant
    f_CodeField = CodeField
    f_NameField = NameField
        
    gDefaSearchCol = DefaultSearchCol
    
    gGridCols = m_Gridcols
    gGridFontName = m_GridFontName
    gGridFontSize = m_FontSize
    ''''''''''''''''''''''''''''''''''''''''''
    
    txtCode.BackColor = m_CodeGotFocusColor
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode)
    
End Sub

Private Sub txtCode_LostFocus()
    txtCode.BackColor = m_CodeLostFocusColor

    f_TabelName = m_TableName
    f_FieldList = m_FieldList
    f_CodeField = m_CodeField
    f_NameField = m_NameField
    f_TopN = m_TopN
    f_SqlWhere = m_SqlWhere
    f_SqlSelect = m_SqlSelect

End Sub

Private Sub txtCode_Validate(Cancel As Boolean)

    If Val(txtCode.Text) > 0 Then
    Dim SQL  As String
    Dim rsttemp As New ADODB.Recordset
        
    If gAdoConnStr = "" Then
        gAdoConnStr = gSetAdoDSN
    End If
        
        If Val(txtCode.Text) <= 0 Then
            rsttemp.Open xSource, gAdoConnStr, adOpenStatic, adLockOptimistic
        Else
            SQL = " Select "
            If Val(m_TopN) > 0 Then
                SQL = SQL & " Top " & TopN
            End If
            SQL = SQL & " * "
            SQL = SQL & " from " & m_TableName
            SQL = SQL & " Where 1=1 "
            SQL = SQL & " And "
            SQL = SQL & m_CodeField & " = " & Val(txtCode.Text)
    
            If Len(m_SqlWhere) > 0 Then
                SQL = SQL & " And " & m_SqlWhere
            End If
    
            rsttemp.Open SQL, gAdoConnStr, adOpenDynamic, adLockOptimistic
            If Not rsttemp.BOF And Not rsttemp.EOF Then
                txtCode = rsttemp.Fields(f_CodeField)
                txtName = rsttemp.Fields(f_NameField)
                Exit Sub
            Else
                If rsttemp.State = adStateOpen Then rsttemp.Close
                rsttemp.Open xSource, gAdoConnStr, adOpenStatic, adLockOptimistic
            End If
        End If
        
        txtCode = ""
        txtName = ""
        frmHlp.msfHelp.Clear
        Set frmHlp.msfHelp.Recordset = rsttemp
        frmHlp.SetMsfDetail frmHlp.msfHelp
    
        If rsttemp.State = adStateOpen Then rsttemp.Close
        Set rsttemp = Nothing
    
        frmHlp.Show vbModal
        If Val(F_Code) <= 0 Then Cancel = True
    Else
        If m_CodeValidate Then
            Cancel = True
        End If
    End If
End Sub

Private Sub txtName_GotFocus()
    If Val(txtCode.Text) <= 0 Then
        txtCode = F_Code
        txtName = F_Name
        F_Code = 0
        F_Name = ""
    End If
    SendKeys "{TAB}"
End Sub

Private Sub UserControl_Resize()

    txtCode.Width = m_CodeWidth
    txtName.Width = m_NameWidth
    'txtName.Visible = m_NameVisible
    
    txtCode.Move 0, 0
    cmdValidate.Move txtCode.Width, 0
    UserControl.Height = txtCode.Height
    
    If m_NameVisible = True Then
        UserControl.Width = txtCode.Width + SpcBtwn + cmdValidate.Width + txtName.Width
        txtName.Move txtCode.Width + SpcBtwn + cmdValidate.Width, 0
    Else
        UserControl.Width = txtCode.Width + cmdValidate.Width
        txtName.Move txtCode.Width, 0
    End If
    
End Sub

Public Function xSource() As String
    Dim SQL As String
    
    If Len(m_SqlSelect) <= 0 Then
        SQL = "Select "
        
        If Val(m_TopN) > 0 Then
            SQL = SQL & " Top " & m_TopN & " "
        End If
        
        SQL = SQL & m_FieldList
        SQL = SQL & " from " & m_TableName
        SQL = SQL & " Where 1=1 "
        If Len(m_SqlWhere) > 0 Then
            SQL = SQL & " And " & m_SqlWhere
        End If
    Else
        SQL = m_SqlSelect
    End If
    
    xSource = SQL
End Function

Public Sub FillFieldCombo()
    MsgBox "Not Supported"
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SqlSelect() As String
    SqlSelect = m_SqlSelect
End Property

Public Property Let SqlSelect(ByVal New_SqlSelect As String)
    m_SqlSelect = New_SqlSelect
    PropertyChanged "SqlSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SqlWhere() As String
    SqlWhere = m_SqlWhere
End Property

Public Property Let SqlWhere(ByVal New_SqlWhere As String)
    m_SqlWhere = New_SqlWhere
    PropertyChanged "SqlWhere"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FieldList() As String
    FieldList = m_FieldList
End Property

Public Property Let FieldList(ByVal New_FieldList As String)
    m_FieldList = New_FieldList
    PropertyChanged "FieldList"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Let TableName(ByVal New_TableName As String)
    m_TableName = New_TableName
    PropertyChanged "TableName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SqlSelect = m_def_SqlSelect
    m_SqlWhere = m_def_SqlWhere
    m_FieldList = m_def_FieldList
    m_TableName = m_def_TableName
    m_TopN = m_def_Topn
    m_GetCode = m_def_GetCode
    m_GetName = m_def_GetName
    m_TopN = m_def_Topn
    m_CodeWidth = m_def_CodeWidth
    m_NameVisible = m_def_NameVisible
    m_SpcBtwn = m_def_SpcBtwn
    m_NameWidth = m_def_NameWidth
    m_NameVisible = m_def_NameVisible
    m_CodeGotFocusColor = m_def_CodeGotFocusColor
    m_CodeLostFocusColor = m_def_CodeLostFocusColor
    m_SetAdoConnStr = m_def_SetAdoConnStr
    m_SetAdoDSN = m_def_SetAdoDSN
    m_CodeField = m_def_CodeField
    m_NameField = m_def_NameField
    m_CodeValidate = m_def_CodeValidate

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_SqlSelect = PropBag.ReadProperty("SqlSelect", m_def_SqlSelect)
    m_SqlWhere = PropBag.ReadProperty("SqlWhere", m_def_SqlWhere)
    m_FieldList = PropBag.ReadProperty("FieldList", m_def_FieldList)
    m_TableName = PropBag.ReadProperty("TableName", m_def_TableName)
    m_TopN = PropBag.ReadProperty("Topn", m_def_Topn)
    m_GetCode = PropBag.ReadProperty("GetCode", m_def_GetCode)
    m_GetName = PropBag.ReadProperty("GetName", m_def_GetName)
    m_TopN = PropBag.ReadProperty("Topn", m_def_Topn)
    m_CodeWidth = PropBag.ReadProperty("CodeWidth", m_def_CodeWidth)
    m_NameVisible = PropBag.ReadProperty("NameVisible", m_def_NameVisible)
    m_SpcBtwn = PropBag.ReadProperty("SpcBtwn", m_def_SpcBtwn)
    m_NameWidth = PropBag.ReadProperty("NameWidth", m_def_NameWidth)
    m_NameVisible = PropBag.ReadProperty("NameVisible", m_def_NameVisible)
    txtCode.Appearance = PropBag.ReadProperty("txtCode3D", 1)
    txtName.Appearance = PropBag.ReadProperty("txtName3D", 1)
    txtCode.BorderStyle = PropBag.ReadProperty("txtCodeBorder", 1)
    txtName.BorderStyle = PropBag.ReadProperty("txtNameBorder", 1)
    txtCode.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    m_CodeGotFocusColor = PropBag.ReadProperty("CodeGotFocusColor", m_def_CodeGotFocusColor)
    m_CodeLostFocusColor = PropBag.ReadProperty("CodeLostFocusColor", m_def_CodeLostFocusColor)
    m_SetAdoConnStr = PropBag.ReadProperty("SetAdoConnStr", m_def_SetAdoConnStr)
    m_SetAdoDSN = PropBag.ReadProperty("SetAdoConnStr", m_def_SetAdoDSN)
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_CodeField = PropBag.ReadProperty("CodeField", m_def_CodeField)
    m_NameField = PropBag.ReadProperty("NameField", m_def_NameField)
    m_CodeValidate = PropBag.ReadProperty("CodeValidate", m_def_CodeValidate)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SqlSelect", m_SqlSelect, m_def_SqlSelect)
    Call PropBag.WriteProperty("SqlWhere", m_SqlWhere, m_def_SqlWhere)
    Call PropBag.WriteProperty("FieldList", m_FieldList, m_def_FieldList)
    Call PropBag.WriteProperty("TableName", m_TableName, m_def_TableName)
    Call PropBag.WriteProperty("GetCode", m_GetCode, m_def_GetCode)
    Call PropBag.WriteProperty("GetName", m_GetName, m_def_GetName)
    Call PropBag.WriteProperty("Topn", m_TopN, m_def_Topn)
    Call PropBag.WriteProperty("CodeWidth", m_CodeWidth, m_def_CodeWidth)
    Call PropBag.WriteProperty("SpcBtwn", m_SpcBtwn, m_def_SpcBtwn)
    Call PropBag.WriteProperty("NameWidth", m_NameWidth, m_def_NameWidth)
    Call PropBag.WriteProperty("NameVisible", m_NameVisible, m_def_NameVisible)
    Call PropBag.WriteProperty("txtCode3D", txtCode.Appearance, 1)
    Call PropBag.WriteProperty("txtName3D", txtName.Appearance, 1)
    Call PropBag.WriteProperty("txtCodeBorder", txtCode.BorderStyle, 1)
    Call PropBag.WriteProperty("txtNameBorder", txtName.BorderStyle, 1)
    Call PropBag.WriteProperty("MaxLength", txtCode.MaxLength, 0)
    Call PropBag.WriteProperty("CodeGotFocusColor", m_CodeGotFocusColor, m_def_CodeGotFocusColor)
    Call PropBag.WriteProperty("CodeLostFocusColor", m_CodeLostFocusColor, m_def_CodeLostFocusColor)
    Call PropBag.WriteProperty("SetAdoConnStr", m_SetAdoConnStr, m_def_SetAdoConnStr)
    Call PropBag.WriteProperty("SetAdoDSN", m_SetAdoDSN, m_def_SetAdoDSN)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("CodeField", m_CodeField, m_def_CodeField)
    Call PropBag.WriteProperty("NameField", m_NameField, m_def_NameField)
    Call PropBag.WriteProperty("CodeValidate", m_CodeValidate, m_def_CodeValidate)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get GetCode() As String
    GetCode = m_GetCode
End Property

Public Property Let GetCode(ByVal New_GetCode As String)
    m_GetCode = New_GetCode
    PropertyChanged "GetCode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get GetName() As String
    GetName = m_GetName
End Property

Public Property Let GetName(ByVal New_GetName As String)
    m_GetName = New_GetName
    PropertyChanged "GetName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get TopN() As Integer
    TopN = m_TopN
End Property

Public Property Let TopN(ByVal New_TopN As Integer)
    m_TopN = New_TopN
    PropertyChanged "Topn"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1200
Public Property Get CodeWidth() As Integer
    CodeWidth = m_CodeWidth
End Property

Public Property Let CodeWidth(ByVal New_CodeWidth As Integer)
    m_CodeWidth = New_CodeWidth
    UserControl_Resize
    PropertyChanged "CodeWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,20
Public Property Get SpcBtwn() As Integer
    SpcBtwn = m_SpcBtwn
End Property

Public Property Let SpcBtwn(ByVal New_SpcBtwn As Integer)
    m_SpcBtwn = New_SpcBtwn
    UserControl_Resize
    PropertyChanged "SpcBtwn"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,3000
Public Property Get NameWidth() As Integer
    NameWidth = m_NameWidth
End Property

Public Property Let NameWidth(ByVal New_NameWidth As Integer)
    m_NameWidth = New_NameWidth
    UserControl_Resize
    PropertyChanged "NameWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get NameVisible() As Boolean
    NameVisible = m_NameVisible
End Property

Public Property Let NameVisible(ByVal New_NameVisible As Boolean)
    m_NameVisible = New_NameVisible
    UserControl_Resize
    PropertyChanged "NameVisible"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,Appearance
Public Property Get txtCode3D() As AppearanceSettings
Attribute txtCode3D.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    txtCode3D = txtCode.Appearance
End Property

Public Property Let txtCode3D(ByVal New_txtCode3D As AppearanceSettings)
    txtCode.Appearance() = New_txtCode3D
    PropertyChanged "txtCode3D"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtName,txtName,-1,Appearance
Public Property Get txtName3D() As AppearanceSettings
Attribute txtName3D.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    txtName3D = txtName.Appearance
End Property

Public Property Let txtName3D(ByVal New_txtName3D As AppearanceSettings)
    txtName.Appearance() = New_txtName3D
    PropertyChanged "txtName3D"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,BorderStyle
Public Property Get txtCodeBorder() As BorderStyleSettings
Attribute txtCodeBorder.VB_Description = "Returns/sets the border style for an object."
    txtCodeBorder = txtCode.BorderStyle
End Property

Public Property Let txtCodeBorder(ByVal New_txtCodeBorder As BorderStyleSettings)
    txtCode.BorderStyle() = New_txtCodeBorder
    PropertyChanged "txtCodeBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtName,txtName,-1,BorderStyle
Public Property Get txtNameBorder() As BorderStyleSettings
Attribute txtNameBorder.VB_Description = "Returns/sets the border style for an object."
    txtNameBorder = txtName.BorderStyle
End Property

Public Property Let txtNameBorder(ByVal New_txtNameBorder As BorderStyleSettings)
    txtName.BorderStyle() = New_txtNameBorder
    PropertyChanged "txtNameBorder"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtCode.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtCode.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CodeGotFocusColor() As OLE_COLOR
    CodeGotFocusColor = m_CodeGotFocusColor
End Property

Public Property Let CodeGotFocusColor(ByVal New_CodeGotFocusColor As OLE_COLOR)
    m_CodeGotFocusColor = New_CodeGotFocusColor
    PropertyChanged "CodeGotFocusColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CodeLostFocusColor() As OLE_COLOR
    CodeLostFocusColor = m_CodeLostFocusColor
End Property

Public Property Let CodeLostFocusColor(ByVal New_CodeLostFocusColor As OLE_COLOR)
    m_CodeLostFocusColor = New_CodeLostFocusColor
    PropertyChanged "CodeLostFocusColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SetAdoConnStr() As Variant
    SetAdoConnStr = m_SetAdoConnStr
End Property

Public Property Let SetAdoConnStr(ByVal New_SetAdoConnStr As Variant)
    m_SetAdoConnStr = New_SetAdoConnStr
    gAdoConnStr = m_SetAdoConnStr
    PropertyChanged "SetAdoConnStr"
End Property

Public Property Get SetAdoDSN() As Variant
    SetAdoDSN = m_SetAdoDSN
End Property

Public Property Let SetAdoDSN(ByVal New_SetAdoDSN As Variant)
    m_SetAdoConnStr = New_SetAdoDSN
    gSetAdoDSN = New_SetAdoDSN
    PropertyChanged "SetAdoDSN"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get CodeText() As Variant
    CodeText = txtCode.Text
End Property

Public Property Let CodeText(ByVal vNewValue As Variant)
    txtCode.Text = vNewValue
    PropertyChanged "CodeText"
End Property

Public Property Get NameText() As Variant
    NameText = txtName.Text
End Property

Public Property Let NameText(ByVal vNewValue As Variant)
    txtName.Text = vNewValue
    PropertyChanged "NameText"
End Property

Public Sub GetNameText(s_Code As Long)
    Dim rsttmp As New ADODB.Recordset
    Dim SQL As String
    
    txtCode.Font.Name = m_CodeFontName
    txtName.Font.Name = m_NameFontName
    
    If s_Code <= 0 Then Exit Sub
    
    f_CodeField = CodeField
    f_NameField = NameField
    
    If gAdoConnStr = "" Then
        gAdoConnStr = gSetAdoDSN
    End If
    
    
    SQL = " Select isnull(" & f_NameField & ",'') as RetVal"
    SQL = SQL & " From " & m_TableName
    SQL = SQL & " Where 1=1 "
    SQL = SQL & " And " & f_CodeField & " = " & Val(s_Code)
    'MsgBox SQL
    rsttmp.Open SQL, gAdoConnStr, adOpenStatic, adLockOptimistic
    
    If rsttmp.RecordCount > 0 Then
        'MsgBox rsttmp.Fields("retVal").Value
        txtName.Text = rsttmp.Fields("retVal").Value
    Else
        txtName.Text = ""
    End If
    
    rsttmp.Close
    Set rsttmp = Nothing
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CodeField() As Variant
    CodeField = m_CodeField
End Property

Public Property Let CodeField(ByVal New_CodeField As Variant)
    m_CodeField = New_CodeField
    f_CodeField = New_CodeField
    PropertyChanged "CodeField"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get NameField() As Variant
    NameField = m_NameField
End Property

Public Property Let NameField(ByVal New_NameField As Variant)
    m_NameField = New_NameField
    f_NameField = New_NameField
    PropertyChanged "NameField"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get CodeValidate() As Boolean
    CodeValidate = m_CodeValidate
End Property

Public Property Let CodeValidate(ByVal New_CodeValidate As Boolean)
    m_CodeValidate = New_CodeValidate
    PropertyChanged "CodeValidate"
End Property

Public Sub ShowHelp()
    cmdValidate_Click
    SendKeys "{TAB}"
    DoEvents
    SendKeys "{TAB}"
    DoEvents
End Sub

Public Sub SetFontParameters(s_CodeFontName As String, s_NameFontName As String, s_GridFontName As String, s_GridCols As String, Optional s_FontSize As Integer = 10)

    If Trim$(s_CodeFontName) = "" Then s_CodeFontName = "Arial"
    If Trim$(s_NameFontName) = "" Then s_NameFontName = "Arial"
    If Trim$(s_GridFontName) = "" Then s_GridFontName = "Arial"
    
    m_CodeFontName = s_CodeFontName
    m_NameFontName = s_NameFontName
    m_GridFontName = s_GridFontName
    m_Gridcols = s_GridCols
    m_FontSize = s_FontSize
    
    gGridCols = s_GridCols
    gGridFontName = s_GridFontName
    gGridFontSize = s_FontSize
    
End Sub

Public Property Get DefaultSearchCol() As Integer
    DefaultSearchCol = gDefaSearchCol
End Property

Public Property Let DefaultSearchCol(ByVal vNewValue As Integer)
    gDefaSearchCol = vNewValue
    PropertyChanged ("DefaultSearchCol")
End Property

Public Function TextMatrixData() As String
    MsgBox "Not Supported"
End Function

Public Property Get TextMatrixDataCol() As Integer
    MsgBox "Not Supported"
End Property

Public Property Let TextMatrixDataCol(ByVal vNewValue As Integer)
    MsgBox "Not Supported"
End Property

Public Function GetFieldValue(ByVal s_FieldName As String, ByVal s_Code As Long) As String
    Dim rsttmp As New ADODB.Recordset
    Dim SQL As String
    
    If s_Code <= 0 Then Exit Function
    
    f_CodeField = CodeField
    f_NameField = NameField
    
    If gAdoConnStr = "" Then
        gAdoConnStr = gSetAdoDSN
    End If
    
    SQL = " Select isnull(" & s_FieldName & ",'') as RetVal"
    SQL = SQL & " From " & m_TableName
    SQL = SQL & " Where 1=1 "
    SQL = SQL & " And " & f_CodeField & " = " & Val(s_Code)
    
    rsttmp.Open SQL, gAdoConnStr, adOpenStatic, adLockOptimistic
    
    If rsttmp.RecordCount > 0 Then
        GetFieldValue = rsttmp.Fields("retVal").Value
    Else
        GetFieldValue = ""
    End If
    
    rsttmp.Close
    Set rsttmp = Nothing
End Function
