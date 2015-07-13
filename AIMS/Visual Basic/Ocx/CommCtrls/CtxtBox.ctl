VERSION 5.00
Begin VB.UserControl CtxtBox 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ScaleHeight     =   2760
   ScaleWidth      =   4995
   Begin VB.TextBox txtC 
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "CtxtBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_AllowNull = False
Const m_def_AutoCaps = True
Const m_def_AdditionalChar = ""
Const m_def_GotFocusClr = 16047579
Const m_def_LostFocusColor = 16777215
'Property Variables:
Dim m_ForeColor As Long
Dim m_AllowNull As Boolean
Dim m_AutoCaps As Boolean
Dim m_AdditionalChar As String
Dim m_GotFocusClr As OLE_COLOR
Dim m_LostFocusColor As OLE_COLOR
'Event Declarations:
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
'Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtC,txtC,-1,KeyPress
Event Validate(Cancel As Boolean) 'MappingInfo=txtC,txtC,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."
Event Change() 'MappingInfo=txtc,txtc,-1,Change
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp

Private Sub txtc_GotFocus()
'16047579 - blue
'16777215 - white
    txtC.BackColor = m_GotFocusClr
    txtC.SelStart = 0
    txtC.SelLength = Len(txtC)
End Sub

Private Sub txtc_LostFocus()
    txtC.BackColor = m_LostFocusColor
End Sub

Private Sub UserControl_Resize()
    txtC.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_GotFocusClr = PropBag.ReadProperty("GotFocusClr", m_def_GotFocusClr)
    m_LostFocusColor = PropBag.ReadProperty("LostFocusColor", m_def_LostFocusColor)
    txtC.Locked = PropBag.ReadProperty("Locked", False)
    txtC.Enabled = PropBag.ReadProperty("Enabled", True)
    txtC.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txtC.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtC.Text = PropBag.ReadProperty("Text", "")
    txtC.Alignment = PropBag.ReadProperty("Alignment", 1)
    txtC.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtC.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtC.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtC.SelLength = PropBag.ReadProperty("SelLength", 0)

    m_AdditionalChar = PropBag.ReadProperty("AdditionalChar", m_def_AdditionalChar)
    m_AutoCaps = PropBag.ReadProperty("AutoCaps", m_def_AutoCaps)
    Set txtC.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_AllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GotFocusClr", m_GotFocusClr, m_def_GotFocusClr)
    Call PropBag.WriteProperty("LostFocusColor", m_LostFocusColor, m_def_LostFocusColor)
    Call PropBag.WriteProperty("Locked", txtC.Locked, False)
    Call PropBag.WriteProperty("Enabled", txtC.Enabled, True)
    Call PropBag.WriteProperty("BackColor", txtC.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderStyle", txtC.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", txtC.Text, "")
    Call PropBag.WriteProperty("Alignment", txtC.Alignment, 1)
    Call PropBag.WriteProperty("Appearance", txtC.Appearance, 1)
    Call PropBag.WriteProperty("MaxLength", txtC.MaxLength, 0)
    Call PropBag.WriteProperty("SelStart", txtC.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtC.SelLength, 0)

    Call PropBag.WriteProperty("AdditionalChar", m_AdditionalChar, m_def_AdditionalChar)
    Call PropBag.WriteProperty("AutoCaps", m_AutoCaps, m_def_AutoCaps)
    Call PropBag.WriteProperty("Font", txtC.Font, Ambient.Font)
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
    m_AutoCaps = m_def_AutoCaps
    m_ForeColor = m_def_ForeColor
    m_AllowNull = m_def_AllowNull
End Sub

Private Sub txtc_Change()
    If m_AutoCaps Then
        OnLineCaps
    End If
    RaiseEvent Change
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
'MappingInfo=txtc,txtc,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txtC.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtC.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = txtC.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtC.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txtC.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtC.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConst
    BorderStyle = txtC.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConst)
    txtC.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,Text
Public Property Get Text() As String
    Text = txtC.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtC.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
    Alignment = txtC.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtC.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
    Appearance = txtC.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    txtC.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txtC.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtC.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txtC.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtC.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtc,txtc,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtC.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtC.SelLength() = New_SelLength
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

Private Sub OnLineCaps()
    Dim i As Integer
    i = txtC.SelStart
    txtC.Text = UCase(txtC.Text)
    txtC.SelStart = i
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoCaps() As Boolean
    AutoCaps = m_AutoCaps
End Property

Public Property Let AutoCaps(ByVal New_AutoCaps As Boolean)
    m_AutoCaps = New_AutoCaps
    PropertyChanged "AutoCaps"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtC,txtC,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtC.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtC.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub txtC_Validate(Cancel As Boolean)
    If Not AllowNull Then
        If Len(txtC.Text) <= 0 Then
            Cancel = True
            txtC.SetFocus
        End If
    End If
    RaiseEvent Validate(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
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

