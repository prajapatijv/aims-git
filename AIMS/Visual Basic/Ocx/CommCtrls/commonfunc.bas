Attribute VB_Name = "Module1"

Public Sub GotFocusEvent(t As TextBox, f_clr As OLE_COLOR)
    t.BackColor = f_clr
    t.SelStart = 0
    t.SelLength = Len(t.Text)
End Sub
