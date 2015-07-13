VERSION 5.00
Begin VB.UserControl frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H008080FF&
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   ScaleHeight     =   4020
   ScaleWidth      =   10440
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
    
    
    With Adodc1
       ' .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=fassoni;Data Source=win98se"
        '.RecordSource =
       ' Set msfHelp.DataSource = Adodc1
    End With
    

End Sub
