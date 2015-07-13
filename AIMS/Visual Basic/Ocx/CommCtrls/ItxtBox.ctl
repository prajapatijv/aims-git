VERSION 5.00
Begin VB.UserControl ItxtBox 
   Alignable       =   -1  'True
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   ScaleHeight     =   2610
   ScaleWidth      =   4545
   Begin VB.TextBox txti 
      Alignment       =   1  'Right Justify
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
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "ItxtBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_MinVal = 0
Const m_def_MaxVal = 32000
Const m_def_AllowNull = False
Const m_def_GotFocusClr = 16047579
Const m_def_LostFocusColor = 16777215

'Property Variables:
Dim m_MinVal As Integer
Dim m_MaxVal As Integer
Dim m_AllowNull As Boolean
Dim m_GotFocusClr As OLE_COLOR
Dim m_LostFocusColor As OLE_COLOR

'Event Declarations:
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event Validate(Cancel As Boolean) 'MappingInfo=txti,txti,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."

'Event KeyPress(KeyAscii As Integer) 'MappingInfo=txti,txti,-1,KeyPress
Event Change() 'MappingInfo=txti,txti,-1,Change
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Enum BorderStyleConst
    vbNone = 0
    vbSingle = 1
End Enum
Private Sub txti_GotFocus()
'16047579 - blue
'16777215 - white
    txti.BackColor = m_GotFocusClr
    txti.SelStart = 0
    txti.SelLength = Len(txti)
End Sub
'
Private Sub txti_LostFocus()
    txti.BackColor = m_LostFocusColor
End Sub

Private Sub UserControl_Resize()
    txti.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_GotFocusClr = PropBag.ReadProperty("GotFocusClr", m_def_GotFocusClr)
    m_LostFocusColor = PropBag.ReadProperty("LostFocusColor", m_def_LostFocusColor)
    txti.Locked = PropBag.ReadProperty("Locked", False)
    txti.Enabled = PropBag.ReadProperty("Enabled", True)
    txti.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txti.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txti.Text = PropBag.ReadProperty("Text", "")
    txti.Alignment = PropBag.ReadProperty("Alignment", 1)
    txti.Appearance = PropBag.ReadProperty("Appearance", 1)
    txti.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txti.SelStart = PropBag.ReadProperty("SelStart", 0)
    txti.SelLength = PropBag.ReadProperty("SelLength", 0)

    Set txti.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txti.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    m_MinVal = PropBag.ReadProperty("MinVal", m_def_MinVal)
    m_MaxVal = PropBag.ReadProperty("MaxVal", m_def_MaxVal)
    m_AllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GotFocusClr", m_GotFocusClr, m_def_GotFocusClr)
    Call PropBag.WriteProperty("LostFocusColor", m_LostFocusColor, m_def_LostFocusColor)
    Call PropBag.WriteProperty("Locked", txti.Locked, False)
    Call PropBag.WriteProperty("Enabled", txti.Enabled, True)
    Call PropBag.WriteProperty("BackColor", txti.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderStyle", txti.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", txti.Text, "")
    Call PropBag.WriteProperty("Alignment", txti.Alignment, 1)
    Call PropBag.WriteProperty("Appearance", txti.Appearance, 1)
    Call PropBag.WriteProperty("MaxLength", txti.MaxLength, 0)
    Call PropBag.WriteProperty("SelStart", txti.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txti.SelLength, 0)

    Call PropBag.WriteProperty("Font", txti.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", txti.ForeColor, &H80000008)
    Call PropBag.WriteProperty("MinVal", m_MinVal, m_def_MinVal)
    Call PropBag.WriteProperty("MaxVal", m_MaxVal, m_def_MaxVal)
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
    m_MinVal = m_def_MinVal
    m_MaxVal = m_def_MaxVal
    m_AllowNull = m_def_AllowNull
End Sub

Private Sub txti_Change()
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
'MappingInfo=txti,txti,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txti.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txti.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = txti.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txti.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txti.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txti.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConst
    BorderStyle = txti.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConst)
    txti.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,Text
Public Property Get Text() As String
    Text = txti.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txti.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
    Alignment = txti.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txti.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
    Appearance = txti.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    txti.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txti.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txti.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txti.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txti.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txti.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txti.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txti.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txti.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txti,txti,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txti.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txti.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MinVal() As Integer
    MinVal = m_MinVal
End Property

Public Property Let MinVal(ByVal New_MinVal As Integer)
    m_MinVal = New_MinVal
    PropertyChanged "MinVal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,32000
Public Property Get MaxVal() As Integer
    MaxVal = m_MaxVal
End Property

Public Property Let MaxVal(ByVal New_MaxVal As Integer)
    m_MaxVal = New_MaxVal
    PropertyChanged "MaxVal"
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

Private Sub txti_Validate(Cancel As Boolean)
    If Not AllowNull Then
        If Not Val(txti.Text) <> 0 Then
            Cancel = True
            txti.SetFocus
        End If
        If Val(txti.Text) < MinVal Or Val(txti.Text) > MaxVal Then
            Cancel = True
            txti.SetFocus
        End If
    End If
    RaiseEvent Validate(Cancel)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

