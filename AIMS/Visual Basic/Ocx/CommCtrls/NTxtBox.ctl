VERSION 5.00
Begin VB.UserControl NTxtBox 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   2760
   ScaleWidth      =   5040
   Begin VB.TextBox txtNtxt 
      Alignment       =   1  'Right Justify
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
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "NTxtBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_MinVal = 0
Const m_def_MaxVal = 99999999
Const m_def_AllowNull = False
Const m_def_Decimal = 2
Const m_def_GotFocusColor = 16047579
Const m_def_LostFocusColor = 16777215

'Property Variables:
Dim m_MinVal As Long
Dim m_MaxVal As Double
Dim m_AllowNull As Boolean
Dim m_Decimal As Integer
Dim m_GotFocusColor As OLE_COLOR
Dim m_LostFocusColor As OLE_COLOR

'Event Declarations:
Event Validate(Cancel As Boolean) 'MappingInfo=txtNtxt,txtNtxt,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event Change() 'MappingInfo=txtNtxt,txtNtxt,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."

Private Sub txtNtxt_GotFocus()
'16047579 - blue
'16777215 - white
    txtNtxt.BackColor = m_GotFocusColor
    txtNtxt.SelStart = 0
    txtNtxt.SelLength = Len(txtNtxt)
End Sub

Private Sub OnlineNums(txt As TextBox)
    If IsNumeric(txt.Text) = False Then
        txt.Text = 0
        txt.SelStart = Len(txt.Text)
    End If
    
    If InStr(1, txt.Text, ".") <> 0 Then
        If txt.MaxLength = 0 Then
            txt.MaxLength = Len(txt.Text) + m_Decimal
        End If
    Else
        txt.MaxLength = 0
    End If
End Sub

Private Sub txtNtxt_LostFocus()
    txtNtxt.BackColor = m_LostFocusColor
    Dim Fmt As String
    Fmt = "########0." & String(m_Decimal, "0")
    txtNtxt.Text = Format(Val(txtNtxt.Text), Fmt)
End Sub

Private Sub UserControl_Resize()
    txtNtxt.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_GotFocusColor = PropBag.ReadProperty("GotFocusColor", m_def_GotFocusColor)
    m_LostFocusColor = PropBag.ReadProperty("LostFocusColor", m_def_LostFocusColor)
    txtNtxt.Locked = PropBag.ReadProperty("Locked", False)
    txtNtxt.Enabled = PropBag.ReadProperty("Enabled", True)
    txtNtxt.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txtNtxt.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtNtxt.Text = PropBag.ReadProperty("Text", "")
    txtNtxt.Alignment = PropBag.ReadProperty("Alignment", 1)
    txtNtxt.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtNtxt.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtNtxt.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtNtxt.SelLength = PropBag.ReadProperty("SelLength", 0)

    m_Decimal = PropBag.ReadProperty("Desimal", m_def_Decimal)
    Set txtNtxt.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtNtxt.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    m_MinVal = PropBag.ReadProperty("MinVal", m_def_MinVal)
    m_MaxVal = PropBag.ReadProperty("MaxVal", m_def_MaxVal)
    m_AllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GotFocusColor", m_GotFocusColor, m_def_GotFocusColor)
    Call PropBag.WriteProperty("LostFocusColor", m_LostFocusColor, m_def_LostFocusColor)
    Call PropBag.WriteProperty("Locked", txtNtxt.Locked, False)
    Call PropBag.WriteProperty("Enabled", txtNtxt.Enabled, True)
    Call PropBag.WriteProperty("BackColor", txtNtxt.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderStyle", txtNtxt.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", txtNtxt.Text, "")
    Call PropBag.WriteProperty("Alignment", txtNtxt.Alignment, 1)
    Call PropBag.WriteProperty("Appearance", txtNtxt.Appearance, 1)
    Call PropBag.WriteProperty("MaxLength", txtNtxt.MaxLength, 0)
    Call PropBag.WriteProperty("SelStart", txtNtxt.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtNtxt.SelLength, 0)

    Call PropBag.WriteProperty("Desimal", m_Decimal, m_def_Decimal)
    Call PropBag.WriteProperty("Font", txtNtxt.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", txtNtxt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("MinVal", m_MinVal, m_def_MinVal)
    Call PropBag.WriteProperty("MaxVal", m_MaxVal, m_def_MaxVal)
    Call PropBag.WriteProperty("AllowNull", m_AllowNull, m_def_AllowNull)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,16047579
Public Property Get GotFocusColor() As OLE_COLOR
    GotFocusColor = m_GotFocusColor
End Property

Public Property Let GotFocusColor(ByVal New_GotFocusColor As OLE_COLOR)
    m_GotFocusColor = New_GotFocusColor
    PropertyChanged "GotFocusColor"
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
    m_GotFocusColor = m_def_GotFocusColor
    m_LostFocusColor = m_def_LostFocusColor
    m_Decimal = m_def_Decimal
    m_MinVal = m_def_MinVal
    m_MaxVal = m_def_MaxVal
    m_AllowNull = m_def_AllowNull
End Sub

Private Sub txtNtxt_Change()
    OnlineNums txtNtxt
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
'MappingInfo=txtNtxt,txtNtxt,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtNtxt.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtNtxt.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtNtxt.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtNtxt.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtNtxt.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtNtxt.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,BorderStyle

Private Property Get BorderStyle() As BorderStyleConst
    BorderStyle = txtNtxt.BorderStyle
End Property

Private Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConst)
    txtNtxt.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtNtxt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtNtxt.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtNtxt.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtNtxt.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = txtNtxt.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    txtNtxt.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtNtxt.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtNtxt.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtNtxt.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtNtxt.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtNtxt.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtNtxt.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get Decimals() As Integer
    Decimals = m_Decimal
End Property

Public Property Let Decimals(ByVal New_Decimals As Integer)
    m_Decimal = New_Decimals
    PropertyChanged "Decimals"
End Property

Private Sub txtNtxt_Validate(Cancel As Boolean)
    If Not AllowNull Then
        If Not Val(txtNtxt.Text) <> 0 Then
            Cancel = True
            txtNtxt.SetFocus
        End If
        If Val(txtNtxt.Text) < MinVal Or Val(txtNtxt.Text) > MaxVal Then
            Cancel = True
            txtNtxt.SetFocus
        End If
    End If

    RaiseEvent Validate(Cancel)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtNtxt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtNtxt.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNtxt,txtNtxt,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtNtxt.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtNtxt.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinVal() As Long
    MinVal = m_MinVal
End Property

Public Property Let MinVal(ByVal New_MinVal As Long)
    m_MinVal = New_MinVal
    PropertyChanged "MinVal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,0,99999999
Public Property Get MaxVal() As Double
    MaxVal = m_MaxVal
End Property

Public Property Let MaxVal(ByVal New_MaxVal As Double)
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

