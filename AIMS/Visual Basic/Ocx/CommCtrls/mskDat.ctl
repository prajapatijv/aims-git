VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl mskDat 
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3585
   ScaleWidth      =   4800
   Begin MSMask.MaskEdBox mskdat 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "mskDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_AllowNull = False
Const m_def_Text = 0

'Property Variables:
Dim m_AllowNull As Boolean
Dim m_Text As Variant

'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=mskdat,mskdat,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=mskdat,mskdat,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=mskdat,mskdat,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Change() 'MappingInfo=mskdat,mskdat,-1,Change
Attribute Change.VB_Description = "Indicates that the contents of a control have changed."
Event Validate(Cancel As Boolean) 'MappingInfo=mskdat,mskdat,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."

Private Sub mskdat_GotFocus()
    mskdat.SelStart = 0
    mskdat.SelLength = Len(mskdat.Text)
End Sub

Private Sub mskdat_Validate(KeepFocus As Boolean)
    If m_AllowNull Then Exit Sub
    KeepFocus = DateValidate
End Sub

Private Sub UserControl_Initialize()
    mskdat.Mask = "##/##/####"
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 1100
    'UserControl.Height = 375
    mskdat.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Private Function DateValidate() As Boolean
    On Error Resume Next
    Dim mDay As Integer
    Dim mMonth As Integer
    Dim mYear As Integer
    Dim mMaxDay As Integer
        
    If mskdat.Text <> "__/__/____" Then
        mDay = Mid(mskdat.Text, 1, 2)
        If mMonth <= 0 Then mMonth = Mid(mskdat.Text, 4, 2)
        If mYear <= 0 Then mYear = Mid(mskdat.Text, 7, 4)
    End If
    
    Select Case mMonth
        Case 1, 3, 5, 7, 8, 10, 12
           mMaxDay = 31
        Case 4, 6, 9, 11
            mMaxDay = 30
        Case 2
            If mYear Mod 4 = 0 Then mMaxDay = 29 Else mMaxDay = 28
    End Select
    
    If mDay > mMaxDay Then
        DateValidate = True
    End If
    
    If mMonth > 12 Then
        DateValidate = True
    End If
            
    If mYear < 1000 Then
        DateValidate = True
    End If
    
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = mskdat.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    mskdat.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = mskdat.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    mskdat.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = mskdat.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    mskdat.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font Property"
Attribute Font.VB_UserMemId = -512
    Set Font = mskdat.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set mskdat.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = mskdat.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    mskdat.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    mskdat.Refresh
End Sub

Private Sub mskdat_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub mskdat_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub mskdat_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Sets whether the control has a flat or sunken 3d appearance"
    Appearance = mskdat.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    mskdat.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,AutoTab
Public Property Get AutoTab() As Boolean
Attribute AutoTab.VB_Description = "Determines whether or not the next control in the tab order receives the focus."
    AutoTab = mskdat.AutoTab
End Property

Public Property Let AutoTab(ByVal New_AutoTab As Boolean)
    mskdat.AutoTab() = New_AutoTab
    PropertyChanged "AutoTab"
End Property

Private Sub mskdat_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mskdat,mskdat,-1,Mask
Public Property Get Mask() As String
Attribute Mask.VB_Description = "Determines the input mask for the control."
    Mask = mskdat.Mask
End Property

Public Property Let Mask(ByVal New_Mask As String)
    mskdat.Mask() = New_Mask
    PropertyChanged "Mask"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AllowNull = m_def_AllowNull
    m_Text = m_def_Text
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    mskdat.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    mskdat.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    mskdat.Enabled = PropBag.ReadProperty("Enabled", True)
    Set mskdat.Font = PropBag.ReadProperty("Font", Ambient.Font)
    mskdat.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    mskdat.Appearance = PropBag.ReadProperty("Appearance", 1)
    mskdat.AutoTab = PropBag.ReadProperty("AutoTab", False)
    mskdat.Mask = PropBag.ReadProperty("Mask", "##/##/####")
    m_AllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", mskdat.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", mskdat.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", mskdat.Enabled, True)
    Call PropBag.WriteProperty("Font", mskdat.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", mskdat.BorderStyle, 1)
    Call PropBag.WriteProperty("Appearance", mskdat.Appearance, 1)
    Call PropBag.WriteProperty("AutoTab", mskdat.AutoTab, False)
    Call PropBag.WriteProperty("Mask", mskdat.Mask, " ##/##/####")
    Call PropBag.WriteProperty("AllowNull", m_AllowNull, m_def_AllowNull)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get AllowNull() As Boolean
    AllowNull = m_AllowNull
End Property

Public Property Let AllowNull(ByVal New_AllowNull As Boolean)
    m_AllowNull = New_AllowNull
    PropertyChanged "AllowNull"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Text() As Variant
    Text = mskdat.Text
End Property

Public Property Let Text(ByVal New_Text As Variant)
    mskdat.Text = New_Text
    PropertyChanged "Text"
End Property

