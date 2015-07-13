VERSION 5.00
Begin VB.UserControl GujTxtBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtCG 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "GUJAFONT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "GUJTxtBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_AllowNull = False
Const m_def_AdditionalChar = ""
Const m_def_GotFocusClr = 16047579
Const m_def_LostFocusColor = 16777215

'Property Variables:
Dim m_ForeColor As Long
Dim m_AllowNull As Boolean
Dim m_AdditionalChar As String
Dim m_GotFocusClr As OLE_COLOR
Dim m_LostFocusColor As OLE_COLOR

'General Variables:
Dim mIsKeyDown As Boolean

'Event Declarations:
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
'Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtCG,txtCG,-1,KeyPress
Event Validate(Cancel As Boolean) 'MappingInfo=txtCG,txtCG,-1,Validate
Event Change() 'MappingInfo=txtCG,txtCG,-1,Change
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp

Private Sub txtCG_GotFocus()
'16047579 - blue
'16777215 - white
    txtCG.BackColor = m_GotFocusClr
    txtCG.SelStart = 0
    txtCG.SelLength = Len(txtCG)
End Sub
'

Private Sub txtCG_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim str As String
    Select Case KeyCode
        Case 96         '0
            str = "î"
        Case 97         '1
            str = "â"
        Case 98         '2
            str = "ä"
        Case 99         '3
            str = "à"
        Case 100        '4
            str = "å"
        Case 101        '5
            str = "ç"
        Case 102        '6
            str = "ê"
        Case 103        '7
            str = "ë"
        Case 104        '8
            str = "è"
        Case 105        '9
            str = "ï"
        Case Else
            Exit Sub
    End Select
    KeyCode = 0
    txtCG.Text = txtCG.Text & Trim(str)
    txtCG.SelStart = Len(txtCG.Text)
    mIsKeyDown = True

End Sub

Private Sub txtCG_KeyPress(KeyAscii As Integer)
    If mIsKeyDown = True Then
        KeyAscii = 0
    End If
    mIsKeyDown = False
End Sub

Private Sub txtCG_LostFocus()
    txtCG.BackColor = m_LostFocusColor
End Sub

Private Sub UserControl_Resize()
    txtCG.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_GotFocusClr = PropBag.ReadProperty("GotFocusClr", m_def_GotFocusClr)
    m_LostFocusColor = PropBag.ReadProperty("LostFocusColor", m_def_LostFocusColor)
    txtCG.Locked = PropBag.ReadProperty("Locked", False)
    txtCG.Enabled = PropBag.ReadProperty("Enabled", True)
    txtCG.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txtCG.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtCG.Text = PropBag.ReadProperty("Text", "")
    txtCG.Alignment = PropBag.ReadProperty("Alignment", 1)
    txtCG.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtCG.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtCG.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtCG.SelLength = PropBag.ReadProperty("SelLength", 0)

    m_AdditionalChar = PropBag.ReadProperty("AdditionalChar", m_def_AdditionalChar)
    Set txtCG.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_AllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GotFocusClr", m_GotFocusClr, m_def_GotFocusClr)
    Call PropBag.WriteProperty("LostFocusColor", m_LostFocusColor, m_def_LostFocusColor)
    Call PropBag.WriteProperty("Locked", txtCG.Locked, False)
    Call PropBag.WriteProperty("Enabled", txtCG.Enabled, True)
    Call PropBag.WriteProperty("BackColor", txtCG.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderStyle", txtCG.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", txtCG.Text, "")
    Call PropBag.WriteProperty("Alignment", txtCG.Alignment, 1)
    Call PropBag.WriteProperty("Appearance", txtCG.Appearance, 1)
    Call PropBag.WriteProperty("MaxLength", txtCG.MaxLength, 0)
    Call PropBag.WriteProperty("SelStart", txtCG.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtCG.SelLength, 0)

    Call PropBag.WriteProperty("AdditionalChar", m_AdditionalChar, m_def_AdditionalChar)
    Call PropBag.WriteProperty("Font", txtCG.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("AllowNull", m_AllowNull, m_def_AllowNull)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,16047579
Public Property Get GotFocusClr() As OLE_COLOR
    GotFocusClr = m_GotFocusClr
End Property

Public Property Let GotFocusClr(ByVal New_GotFocusClr As OLE_COLOR)
    m_GotFocusClr = New_GotFocusClr
    PropertyChanged "GotFocusClr"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,16777215
Public Property Get LostFocusColor() As OLE_COLOR
    LostFocusColor = m_LostFocusColor
End Property

Public Property Let LostFocusColor(ByVal New_LostFocusColor As OLE_COLOR)
    m_LostFocusColor = New_LostFocusColor
    PropertyChanged "LostFocusColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_GotFocusClr = m_def_GotFocusClr
    m_LostFocusColor = m_def_LostFocusColor
    m_AdditionalChar = m_def_AdditionalChar
    m_ForeColor = m_def_ForeColor
    m_AllowNull = m_def_AllowNull
    mIsKeyDown = False
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txtCG.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtCG.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = txtCG.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtCG.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txtCG.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtCG.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConst
    BorderStyle = txtCG.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConst)
    txtCG.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,Text
Public Property Get Text() As String
    Text = txtCG.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtCG.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
    Alignment = txtCG.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtCG.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
    Appearance = txtCG.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    txtCG.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txtCG.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtCG.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txtCG.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtCG.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtCG.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtCG.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get AdditionalChar() As String
    AdditionalChar = m_AdditionalChar
End Property

Public Property Let AdditionalChar(ByVal New_AdditionalChar As String)
    m_AdditionalChar = New_AdditionalChar
    PropertyChanged "AdditionalChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCG,txtCG,-1,Font
Public Property Get Font() As Font
    Set Font = txtCG.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtCG.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub txtCG_Validate(Cancel As Boolean)
    If Not AllowNull Then
        If Len(txtCG.Text) <= 0 Then
            Cancel = True
            txtCG.SetFocus
        End If
    End If
    RaiseEvent Validate(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,FALSE
Public Property Get AllowNull() As Boolean
    AllowNull = m_AllowNull
End Property

Public Property Let AllowNull(ByVal New_AllowNull As Boolean)
    m_AllowNull = New_AllowNull
    PropertyChanged "AllowNull"
End Property

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub


