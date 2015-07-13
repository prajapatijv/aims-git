Attribute VB_Name = "PkgAumsoft"
Option Explicit
Global Const gPkgName As String = "AumInvontry"
Global Const gDemoMode As Boolean = False

Public Enum enUserLevel
    eAdmin = 1
    eImpex = 2
    eTraining = 3
    eUser = 4
End Enum

Public Function SetPkgName() As String
    SetPkgName = gPkgName
End Function

Public Sub chkStructure()
'---1. KeybrdItem / KeybrdSetup
    If adofieldDet("KeybrdItem", "keybrd_code", , gCnnMst) = False Then AlterTableColumn "KeybrdItem", "keybrd_code", "Numeric", 4, , "add", , gCnnMst
    If adofieldDet("KeybrdItem", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "KeybrdItem", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
    If adofieldDet("KeybrdSetup", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "KeybrdSetup", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---2. ServerExport
    If adofieldDet("ServerExport", "TableType", , gCnnMst) = False Then AlterTableColumn "ServerExport", "TableType", "Varchar", 10, , "add", , gCnnMst
    If adofieldDet("ServerExport", "sWhere", , gCnnMst) = False Then AlterTableColumn "ServerExport", "sWhere", "Varchar", 500, , "add", , gCnnMst

'---3. Items
    If adofieldDet("Items", "unit_id", , gCnnMst) = False Then AlterTableColumn "Items", "unit_id", "Numeric", 4, , "add", , gCnnMst
    If adofieldDet("Items", "BillFmt", , gCnnMst) = False Then AlterTableColumn "Items", "BillFmt", "SmallInt", , , "add", , gCnnMst
    If adofieldDet("Items", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Items", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst

'---4. Sales - Saltrn : TerSaltrn
    '---Server
    If adofieldDet("Saltrn", "canceled", , gCnnMst) = False Then AlterTableColumn "Saltrn", "canceled", "Bit", , , "add", , gCnnMst
    If adofieldDet("Saltrn", "Event_id", , gCnnMst) = False Then AlterTableColumn "Saltrn", "Event_id", "SmallInt", , , "add", , gCnnMst
    If adofieldDet("Saltrn", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Saltrn", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    If adofieldDet("Saldet", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Saldet", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
    '---Terminal
    If adofieldDet("TerSaltrn", "canceled", , gCnnMst) = False Then AlterTableColumn "TerSaltrn", "canceled", "Bit", , , "add", "default(0)", gCnnMst
    If adofieldDet("TerSaltrn", "Event_id", , gCnnMst) = False Then AlterTableColumn "TerSaltrn", "Event_id", "SmallInt", , , "add", "Default(0)", gCnnMst
    If adofieldDet("TerSaltrn", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "TerSaltrn", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    If adofieldDet("TerSaldet", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "TerSaldet", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---5. BillSequence
    If adofieldDet("BillSequence", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "BillSequence", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst

'---6. Categories
    If adofieldDet("Categories", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Categories", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst

'---7. EventMast
    If adofieldDet("EventMast", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "EventMast", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---8. Locations
    If adofieldDet("Locations", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Locations", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---9. Sizes
    If adofieldDet("Sizes", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Sizes", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst

'---10.TerminalConfig
    If adofieldDet("TerminalConfig", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "TerminalConfig", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---11.Units
    If adofieldDet("Units", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Units", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---12.UserMast
    If adofieldDet("UserMast", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "UserMast", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst

'---13.Invdet
    If adofieldDet("Invdet", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Invdet", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst
    
'---14.Invtrn
    If adofieldDet("Invtrn", "Trng_fg", , gCnnMst) = False Then AlterTableColumn "Invtrn", "Trng_fg", "Bit", , , "add", "Default(0)", gCnnMst

'    If adofieldDet("PatMast", "PatName", , gCnnMst) Then
'        If adofieldDet("PatMast", "PatName", "length", gCnnMst) <= 30 Then
'            AlterTableColumn "PatMast", "PatName", "Varchar", 60, , "alter", , gCnnMst
'        End If
'    End If

End Sub

Public Function SqlCategoryMast(f_CategoryMast) As String
    SQL = "create table " & f_CategoryMast & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(30)     Not Null"
    SQL = SQL & ",shortname     varchar(30)     Not Null"
    SQL = SQL & ",actv_fg       Bit         default(0)"
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlCategoryMast = SQL
End Function

Public Function SqlLocationMast(f_LocationMast) As String
    SQL = "create table " & f_LocationMast & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(30)     Not Null"
    SQL = SQL & ",shortname     varchar(30)     Not Null"
    SQL = SQL & ",actv_fg       Bit         default(0)"
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlLocationMast = SQL
End Function

Public Function SqlSizeMast(f_SqlSizeMast) As String
    SQL = "create table " & f_SqlSizeMast & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(30)     Not Null"
    SQL = SQL & ",shortname     varchar(30)     Not Null"
    SQL = SQL & ",actv_fg       Bit         default(0)"
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
        
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlSizeMast = SQL
End Function

Public Function SqlUnitMast(f_SqlUnitMast) As String
    SQL = "create table " & f_SqlUnitMast & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(30)     Not Null"
    SQL = SQL & ",shortname     varchar(30)     Not Null"
    SQL = SQL & ",actv_fg       Bit         default(0)"
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlUnitMast = SQL
End Function

Public Function SqlItemMast(f_SqlItemMast) As String
    SQL = "create table " & f_SqlItemMast & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(60)     Not Null"
    SQL = SQL & ",shortname     varchar(35)     Not Null"
    SQL = SQL & ",actv_fg       Bit             default(0)"
    
    SQL = SQL & ",category_id   Numeric(4)"
    SQL = SQL & ",size_id       Numeric(4)"
    
    SQL = SQL & ",rtl_prc       Numeric(10,2)"
    SQL = SQL & ",disc_per      Numeric(5,2)"
    SQL = SQL & ",disc_amt      Numeric(8,2)"

    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",unit_id       Numeric(4)      default(0)"
    SQL = SQL & ",BillFmt       SmallInt"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    
    SqlItemMast = SQL
End Function

Public Function SqlTerminalConfig(f_SqlTerminalConfig) As String
    SQL = "create table " & f_SqlTerminalConfig & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(60)     Not Null"
    SQL = SQL & ",shortname     varchar(35)     Not Null"
    SQL = SQL & ",actv_fg       Bit             default(0)"
    
    SQL = SQL & ",keybrd_id     Numeric(4)"
    
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlTerminalConfig = SQL
End Function

Public Function SqlKeybrdItem(f_SqlKeybrdItem) As String
    
    SQL = "create table " & f_SqlKeybrdItem & "("
    SQL = SQL & " keybrd_code   Numeric(4)      Not Null"
    SQL = SQL & ",seq           smallint        Not Null"
    SQL = SQL & ",itm_code      Numeric(4)      Not Null"
    SQL = SQL & ",actv_fg       Bit             default(0)"
    
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlKeybrdItem = SQL
End Function

Public Function SqlKeybrdSetup(f_SqlKeybrdSetup) As String
    
    SQL = "create table " & f_SqlKeybrdSetup & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(60)     Not Null"
    SQL = SQL & ",actv_fg       Bit             default(0)"
    
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlKeybrdSetup = SQL
End Function

Public Function SqlTerSaltrn(f_SqlTerSaltrn) As String
    
    SQL = "create table " & f_SqlTerSaltrn & "("
    SQL = SQL & " tran_id           Varchar(20)      Not Null"
    SQL = SQL & ",ter_id            smallint         Not Null"
    SQL = SQL & ",export_fg         Bit             default(0)"
    
    SQL = SQL & ",paid_amt          Numeric(10,2)"
    SQL = SQL & ",change_amt        Numeric(7,2)"
    
    SQL = SQL & ",dtadat            datetime"
    SQL = SQL & ",dtatim            varchar(10)"
    SQL = SQL & ",dtausr            varchar(10)     default('')"
    SQL = SQL & ",Canceled          Bit"
    SQL = SQL & ",Event_id          TinyInt         default(0)"
    
    SQL = SQL & ",Trng_fg           Bit             default(0)"

    SQL = SQL & ")"
    SqlTerSaltrn = SQL
End Function

Public Function SqlTerSaldet(f_SqlTerSaldet) As String
    
    SQL = "create table " & f_SqlTerSaldet & "("
    SQL = SQL & " tran_id       Varchar(20)      Not Null"
    SQL = SQL & ",tran_seq      smallint         Not Null"
    SQL = SQL & ",itm_code      Numeric(4)"
    
    SQL = SQL & ",rtl_prc       Numeric(10,2)"
    SQL = SQL & ",disc_amt      Numeric(10,2)"
    SQL = SQL & ",qty           smallint"
    SQL = SQL & ",amt           numeric(10,2)"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlTerSaldet = SQL
End Function

Public Function SqlBillSequence(f_SqlBillSequence) As String
    
    SQL = "create table " & f_SqlBillSequence & "("
    SQL = SQL & " dtadat            varchar(6)      Not Null"
    SQL = SQL & ",ter_id            smallint        Not Null"
    SQL = SQL & ",seq               smallint        default(0)"
    
    SQL = SQL & ",Trng_fg           Bit             default(0)"
    
    SQL = SQL & ")"
    SqlBillSequence = SQL
End Function

Public Function SqlInvtrn(f_Invtrn) As String
    SQL = "create table " & f_Invtrn & "("
    SQL = SQL & " vno           Numeric(6)      Not Null"
    SQL = SQL & ",ter_id        smallint        Not Null"
    SQL = SQL & ",export_fg     Bit             default(0)"
    
    SQL = SQL & ",tran_type     smallint        Not Null"
    '   1   -   Opening Balance
    '   2   -   Material Inward
    '   3   -   Stock Adjustment Up
    '   4   -   Receive from Store
    '   11  -   Stock Adjustment Down
    '   12  -   Stock Waste
    '   13  -   Issue For Sale

    SQL = SQL & ",doc_no        varchar(30)     "
    SQL = SQL & ",rec_dat       Datetime        Not Null"
    
    SQL = SQL & ",remarks       varchar(60)    Null"
    
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlInvtrn = SQL
End Function

Public Function SqlInvdet(f_Invdet) As String
    SQL = "create table " & f_Invdet & "("
    SQL = SQL & " vno           Numeric(6)      Not Null"
    SQL = SQL & ",srno          int             Not Null"
    SQL = SQL & ",itm_code      Numeric(4)      Not Null"
    
    SQL = SQL & ",qty           smallint        Default(0)"
    SQL = SQL & ",rtl_prc       Numeric(10,2)   Default(0)"
    SQL = SQL & ",amt           Numeric(10,2)   Default(0)"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    
    SqlInvdet = SQL
End Function

'''Public Sub ReadComPort()
'''On Error GoTo ErrHandle
'''
'''Dim mCommPortNo As Integer
'''
'''    If len(AdoParaRead("scale_port", , , "paramete")) = 0 Then
'''        MsgBox "Please Define Scale Port Number in paramete ......" & vbCrLf _
'''        & " para  = " & "Scale_port" & vbCrLf _
'''        & " value = 1"
'''        mCommPortNo = 1
'''    Else
'''        mCommPortNo = Val(AdoParaRead("Scale_port", , , "paramete"))
'''    End If
'''
'''    With mdiMainMenu.mscRead_comm
''''        mCommPortNo = IIf(AdoParaRead("scale_port") = "", 1, AdoParaRead("scale_port"))
'''        If Not .PortOpen Then
'''            .CommPort = mCommPortNo
'''            .PortOpen = True
'''        End If
'''        .Inpulen = 0
'''        .InputMode = comInputModeText
'''    End With
'''
'''Exit Sub
'''ErrHandle:
'''    MsgBox err.Description
'''    Resume Next
'''
'''End Sub

Public Function SqlFormMast(f_FormMast) As String
    SQL = "create table " & f_FormMast & "("
    SQL = SQL & " formid         Numeric(4)      Not Null"
    SQL = SQL & ",fname         Numeric(4)      Not Null"
    SQL = SQL & ",fdescription          Numeric(4)      Not Null"
    SQL = SQL & ",fcaption           Numeric(4,2)"
    SQL = SQL & ",fshortname         varchar(15)"
    SQL = SQL & ",fwhattodo          varchar(20)"
    SQL = SQL & ",status        char(1)         default('T')"
    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"
    SQL = SQL & ")"
    SqlFormMast = SQL
End Function

Public Function SqlUserMast(f_UserMast) As String
    SQL = "create table " & f_UserMast & "("
    SQL = SQL & " Uid           Numeric(4)      Not Null"
    SQL = SQL & ",Uname         varchar(15)     Not Null"
    SQL = SQL & ",pwd           varchar(10)     Not Null"
    SQL = SQL & ",Level         Int             Null"
    SQL = SQL & ",status        char(1)         default('T')"
    
    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    SqlUserMast = SQL
End Function

Public Function SqlUserRights(f_UserRights) As String
    SQL = "create table " & f_UserRights & "("
    SQL = SQL & " Uid           Numeric(4)      Not Null"
    SQL = SQL & ",FormId        Numeric(4)      Not Null"
    SQL = SQL & ",[add]         Char(1)         Null"
    SQL = SQL & ",[del]         Char(1)         Null"
    SQL = SQL & ",[edit]        Char(1)         Null"
    SQL = SQL & ",[print]       Char(1)         Null"
    SQL = SQL & ",[Nevigate]    Char(1)         Null"
    SQL = SQL & ",[report]      Char(1)         Null"
    SQL = SQL & ",status        char(1)         default('T')"
    SQL = SQL & ")"
    SqlUserRights = SQL
End Function

Public Function SqlFormCollection(f_FormCollection) As String
    SQL = "create table " & f_FormCollection & "("
    SQL = SQL & " FormId        Numeric(4)      Identity(1,1)"
    SQL = SQL & ",FormName      VarChar(30)     Not Null"
    SQL = SQL & ",FormCap       VarChar(50)     Null"
    SQL = SQL & ",Descr         VarChar(50)     Null"
    SQL = SQL & ",FormGrp       Char(1)         Null" 'M-Master/T-Transction/R-Report
    SQL = SQL & ",status        char(1)         default('T')"
    SQL = SQL & ")"
    SqlFormCollection = SQL
End Function

Public Function SqlServerExport(f_ServerExport) As String
    SQL = "create table " & f_ServerExport & "("
    SQL = SQL & " Id            Numeric(4)      Identity(1,1)"
    SQL = SQL & ",TableName     VarChar(30)     Not Null"
    SQL = SQL & ",TableType     VarChar(10)     Not Null"
    SQL = SQL & ",sWhere        VarChar(500)"
    SQL = SQL & ",actv_fg       Bit"
    SQL = SQL & ")"
    SqlServerExport = SQL
End Function

Public Function SqlDenoms(f_SqlDenoms) As String
    SQL = "create table " & f_SqlDenoms & "("
    SQL = SQL & " denom_id      Numeric(4)"
    SQL = SQL & ",Descr         VarChar(60)"
    SQL = SQL & ",denom_struc   Numeric(4)"
    SQL = SQL & ")"
    SqlDenoms = SQL
End Function

Public Function SqlEventMast(f_SqlEventMast) As String
    SQL = "create table " & f_SqlEventMast & "("
    SQL = SQL & " code          Numeric(4)      Not Null"
    SQL = SQL & ",name          varchar(60)     Not Null"
    SQL = SQL & ",shortname     varchar(35)     Not Null"
    SQL = SQL & ",actv_fg       Bit             default(0)"
    
    SQL = SQL & ",location_id   Numeric(4)"
    SQL = SQL & ",fdat          datetime"
    SQL = SQL & ",tdat          datetime"

    SQL = SQL & ",dtadat        datetime"
    SQL = SQL & ",dtatim        varchar(10)"
    SQL = SQL & ",dtausr        varchar(10)     default('')"

    SQL = SQL & ",Trng_fg       Bit             default(0)"
    
    SQL = SQL & ")"
    
    SqlEventMast = SQL
End Function


Public Sub CreatePkgMstTables()
'---Package Master Tables
    '1.
    '2.
    If Not AdoIsTable("Categories", gCnnMst) Then
        gCnnMst.Execute SqlCategoryMast("Categories")
    End If
    '3.
    If Not AdoIsTable("Locations", gCnnMst) Then
        gCnnMst.Execute SqlLocationMast("Locations")
    End If
    '4.
    If Not AdoIsTable("Sizes", gCnnMst) Then
        gCnnMst.Execute SqlSizeMast("Sizes")
    End If
    '5.
    If Not AdoIsTable("Items", gCnnMst) Then
        gCnnMst.Execute SqlItemMast("Items")
    End If
    '6.
    If Not AdoIsTable("KeybrdItem", gCnnMst) Then
        gCnnMst.Execute SqlKeybrdItem("KeybrdItem")
    End If
    '7.
    If Not AdoIsTable("KeybrdSetup", gCnnMst) Then
        gCnnMst.Execute SqlKeybrdSetup("KeybrdSetup")
    End If
    '8.
    If Not AdoIsTable("ServerExport", gCnnMst) Then
        gCnnMst.Execute SqlServerExport("ServerExport")
        DefaultEntries "ServerExport"
    End If
    '9.
    If Not AdoIsTable("Denoms", gCnnMst) Then
        gCnnMst.Execute SqlDenoms("Denoms")
        DefaultEntries "Denoms"
    End If
    '10
    If Not AdoIsTable("TerminalConfig", gCnnMst) Then
        gCnnMst.Execute SqlTerminalConfig("TerminalConfig")
    End If
    '11.
    If Not AdoIsTable("Units", gCnnMst) Then
        gCnnMst.Execute SqlUnitMast("Units")
    End If
    '12.
    If Not AdoIsTable("EventMast", gCnnMst) Then
        gCnnMst.Execute SqlEventMast("EventMast")
    End If
    

    
'---SetAuthority Tables
    '1.
    If Not AdoIsTable("UserMast", gCnnMst) Then
        gCnnMst.Execute SqlUserMast("UserMast")
        DefaultEntries "UserMast"
    End If
    '2.
    If Not AdoIsTable("FormCollection", gCnnMst) Then
        gCnnMst.Execute SqlFormCollection("FormCollection")
        DefaultEntries "formcollection"
    End If
    '3.
    If Not AdoIsTable("UserRights", gCnnMst) Then
        gCnnMst.Execute SqlUserRights("UserRights")
        DefaultEntries "UserRights"
    End If
    
End Sub

Public Sub CreatePkgTrnTables()
'---Package Transction Table
    
'---Terminal
        '1.
        If Not AdoIsTable("TerSaltrn", gCnnMst) Then
            gCnnMst.Execute SqlTerSaltrn("TerSaltrn")
        End If
        
        '2.
        If Not AdoIsTable("TerSaldet", gCnnMst) Then
            gCnnMst.Execute SqlTerSaldet("TerSaldet")
        End If
        
        '3.
        If Not AdoIsTable("BillSequence", gCnnMst) Then
            gCnnMst.Execute SqlBillSequence("BillSequence")
        End If

'---Server
        '4.
        If Not AdoIsTable("Saltrn", gCnnMst) Then
            gCnnMst.Execute SqlTerSaltrn("Saltrn")
        End If
    
        '5.
        If Not AdoIsTable("Saldet", gCnnMst) Then
            gCnnMst.Execute SqlTerSaldet("Saldet")
        End If
    
        '6.
        If Not AdoIsTable("Invtrn", gCnnMst) Then
            gCnnMst.Execute SqlInvtrn("Invtrn")
        End If
    
        '7.
        If Not AdoIsTable("Invdet", gCnnMst) Then
            gCnnMst.Execute SqlInvdet("Invdet")
        End If
    
End Sub


Public Sub AddPkgConstraintsMst()
'---Master Table Constraint
    '1.
    '2.
    If Not AdoIsConstraint("Pk_Categories", gCnnMst) Then
        SQL = "Alter Table Categories Add Constraint Pk_Categories Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '3.
    If Not AdoIsConstraint("Pk_Locations", gCnnMst) Then
        SQL = "Alter Table Locations Add Constraint Pk_Locations Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '4.
    If Not AdoIsConstraint("Pk_Sizes", gCnnMst) Then
        SQL = "Alter Table Sizes Add Constraint Pk_Sizes Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '5.
    If Not AdoIsConstraint("Pk_Items", gCnnMst) Then
        SQL = "Alter Table Items Add Constraint Pk_Items Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '6.
    If Not AdoIsConstraint("Pk_KeybrdItem", gCnnMst) Then
        SQL = "Alter Table KeybrdItem Add Constraint Pk_KeybrdItem Primary Key(keybrd_code,itm_code,seq)"
        gCnnMst.Execute SQL
    End If
    '7.
    If Not AdoIsConstraint("Pk_KeybrdSetup", gCnnMst) Then
        SQL = "Alter Table KeybrdSetup Add Constraint Pk_KeybrdSetup Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '8. ServerExport
        'No Constarint
        
    '9. Denoms
        'No Constarint
        
    '10.
    If Not AdoIsConstraint("Pk_TerminalConfig", gCnnMst) Then
        SQL = "Alter Table TerminalConfig Add Constraint Pk_TerminalConfig Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '11.
    If Not AdoIsConstraint("Pk_Units", gCnnMst) Then
        SQL = "Alter Table Units Add Constraint Pk_Units Primary Key(Code)"
        gCnnMst.Execute SQL
    End If
    '12.
    If Not AdoIsConstraint("Pk_EventMast", gCnnMst) Then
        SQL = "Alter Table EventMast Add Constraint Pk_EventMast Primary Key(Code)"
        gCnnMst.Execute SQL
    End If

       
'---SetAuthority Tables
    '1.
    If Not AdoIsConstraint("Pk_UserMast", gCnnMst) Then
        SQL = "Alter Table UserMast Add Constraint Pk_UserMast Primary Key(Uname)"
        gCnnMst.Execute SQL
    End If
    '2.
    If Not AdoIsConstraint("Pk_UserRights", gCnnMst) Then
        SQL = "Alter Table UserRights Add Constraint Pk_UserRights Primary Key(Uid,Formid)"
        gCnnMst.Execute SQL
    End If
    '3.
    If Not AdoIsConstraint("Pk_FormCollection", gCnnMst) Then
        SQL = "Alter Table FormCollection Add Constraint Pk_FormCollection Primary Key(FormName)"
        gCnnMst.Execute SQL
    End If
    
End Sub

Public Sub AddPkgConstraintsTrn()
'---Transction Table Constraint
    
'---Terminal
        '1.
        If Not AdoIsConstraint("Pk_TerSaltrn", gCnnMst) Then
            SQL = "Alter Table TerSaltrn Add Constraint Pk_TerSaltrn Primary Key(tran_id)"
            gCnnMst.Execute SQL
        End If
        
        '2.
        If Not AdoIsConstraint("Pk_TerSaldet", gCnnMst) Then
            SQL = "Alter Table TerSaldet Add Constraint Pk_TerSaldet Primary Key(tran_id,tran_seq)"
            gCnnMst.Execute SQL
        End If
        
        '3.
        If Not AdoIsConstraint("Pk_BillSequence", gCnnMst) Then
            SQL = "Alter Table BillSequence Add Constraint Pk_BillSequence Primary Key(dtadat,ter_id)"
            gCnnMst.Execute SQL
        End If
    
'---Server
        '4.
        If Not AdoIsConstraint("Pk_Saltrn", gCnnMst) Then
            SQL = "Alter Table Saltrn Add Constraint Pk_Saltrn Primary Key(tran_id)"
            gCnnMst.Execute SQL
        End If
    
        '5.
        If Not AdoIsConstraint("Pk_Saldet", gCnnMst) Then
            SQL = "Alter Table Saldet Add Constraint Pk_Saldet Primary Key(tran_id,tran_seq)"
            gCnnMst.Execute SQL
        End If
        
        '6.
        If Not AdoIsConstraint("Pk_Invtrn", gCnnMst) Then
            SQL = "Alter Table Invtrn Add Constraint Pk_Invtrn Primary Key(vno)"
            gCnnMst.Execute SQL
        End If
        
        '7.
        If Not AdoIsConstraint("Pk_Invdet", gCnnMst) Then
            SQL = "Alter Table Invdet Add Constraint Pk_Invdet Primary Key(vno,srno)"
            gCnnMst.Execute SQL
        End If

End Sub


Private Sub FormCollectionDetaults()
    '--------------MASTER FORMS----------------
    Add2FormCollection "FRMACMAST", "ACCOUNT MASTER", "ACCOUNT MASTER", "M", "T"
    Add2FormCollection "FRMCOMPMAST", "COMPANY MASTER", "COMPANY MASTER", "M", "T"
    
    Add2FormCollection "FRMCRVIEWER", "CRVIEWER", "CRVIEWER", "M", "F"
    
    '--------------TRANSCTIONS FORMS----------------
    Add2FormCollection "FRMBATCHTRN", "HEAT ENTRY", "HEAT ENTRY", "T", "T"
    
    '--------------REPORTS FORMS----------------
    Add2FormCollection "FRMREPSPROD", "PRODUCTION & CONSUMPTION REPORTS", "PRODUCTION & CONSUMPTION REPORTS", "R", "T"
    
    '--------------OTHER FORMS----------------
    Add2FormCollection "FRMUSERMAST", "USER MASTER", "USER MASTER", "X", "T"
    
End Sub

Private Sub ServerExportDetaults()
    
'---SERVER SIDE TABLES----------------------------------------------------
    '--------------User Tables------------------------
    Add2ServerExport "Categories", "Server", 1
    Add2ServerExport "Locations", "Server", 1
    Add2ServerExport "Sizes", "Server", 1
    Add2ServerExport "Denoms", "Server", 1
    Add2ServerExport "Units", "Server", 1
    Add2ServerExport "EventMast", "Server", 1
    
    Add2ServerExport "Items", "Server", 1
    
    '--------------Configuration Tables----------------
    Add2ServerExport "KeybrdSetup", "Server", 1
    Add2ServerExport "KeybrdItem", "Server", 1
    Add2ServerExport "UserMast", "Server", 1
    Add2ServerExport "TerminalConfig", "Server", 1

'---TERMINAL SIDE TABLES----------------------------------------------------
    '--------------User Tables------------------------
    Add2ServerExport "TerSaltrn", "Terminal", 1
    Add2ServerExport "TerSaldet", "Terminal", 1
    
End Sub

Private Sub DenomDefaults()

    Add2Denoms 1, "Thousand", 1000
    Add2Denoms 2, "Five Hundred", 500
    Add2Denoms 3, "Hundred", 100
    Add2Denoms 4, "Fifty", 50
    Add2Denoms 5, "Twenty", 20
    Add2Denoms 6, "Ten", 10
    Add2Denoms 7, "Five", 5
    Add2Denoms 8, "Two", 2
    Add2Denoms 9, "One", 1
    
End Sub

Public Sub DefaultEntries(s_Mode As String)
    
    Select Case LCase(s_Mode)
        Case "usermast"
            DefaultUsers

        Case "userrights"
            DefaultRights 1001
        
        Case "serverexport"
            ServerExportDetaults
            
        Case "denoms"
            DenomDefaults
        
        Case "formcollection"
            FormCollectionDetaults
            
    End Select
End Sub

Public Sub DefaultRights(s_Uid As Integer)
    Dim rst As ADODB.Recordset
    
    SQL = "select top 1 'True' "
    SQL = SQL & " from UserRights"
    SQL = SQL & " Where 1=1"
    SQL = SQL & " And Uid = " & Val(s_Uid)
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    If rst.RecordCount <= 0 Then
        CloseAdoRst rst, False
        
        SQL = "Select FormId from FormCollection where status = 'T'"
        OpenAdoRst rst, SQL, , , , gCnnMst
        With rst
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    SQL = "Insert into UserRights ("
                    SQL = SQL & "Uid"
                    SQL = SQL & ",Formid"
                    SQL = SQL & ",[Add]"
                    SQL = SQL & ",[Del]"
                    SQL = SQL & ",[Edit]"
                    SQL = SQL & ",[Print]"
                    SQL = SQL & ",[Nevigate]"
                    SQL = SQL & ",[Report]"
                    SQL = SQL & ",[Status]"
                    
                    SQL = SQL & " ) Values ("
                    
                    SQL = SQL & Val(s_Uid)
                    SQL = SQL & "," & Val(.Fields("FormId").Value)
                    SQL = SQL & ",'1'" 'add
                    SQL = SQL & ",'1'" 'del
                    SQL = SQL & ",'1'" 'edit
                    SQL = SQL & ",'1'" 'print
                    SQL = SQL & ",'1'" 'nevigate
                    SQL = SQL & ",'1'" 'report
                    SQL = SQL & ",'T'" 'Status
                    SQL = SQL & ")"
                    gCnnMst.Execute SQL
                    
                    .MoveNext
                Wend
            End If
        End With
        CloseAdoRst rst
    End If
    
End Sub

Private Sub Add2FormCollection(s_FormName As String, s_FormCap As String, s_Descr As String, s_FormGrp As String, s_Status As String)
    Dim rst As ADODB.Recordset
    
    SQL = "Select 'True' from FormCollection"
    SQL = SQL & " Where FormName = " & AQ(s_FormName)
    OpenAdoRst rst, SQL, , , , gCnnMst
    
    If rst.RecordCount <= 0 Then
        CloseAdoRst rst
        
        SQL = "Insert into FormCollection ("
        SQL = SQL & " FormName"
        SQL = SQL & ",FormCap"
        SQL = SQL & ",Descr"
        SQL = SQL & ",FormGrp"
        SQL = SQL & ",Status"
        SQL = SQL & " )values ("
        SQL = SQL & AQ(s_FormName)
        SQL = SQL & "," & AQ(s_FormCap)
        SQL = SQL & "," & AQ(s_Descr)
        SQL = SQL & "," & AQ(s_FormGrp)
        SQL = SQL & "," & AQ(s_Status)
        SQL = SQL & ")"
        
        gCnnMst.Execute SQL
    End If
End Sub

Private Sub Add2ServerExport(s_TableName As String, s_TableType As String, s_ActvFg As Integer)
    
        SQL = "If Not Exists " & vbCrLf
        SQL = SQL & " ( Select '1' from ServerExport "
        SQL = SQL & "  Where TableName = " & AQ(s_TableName)
        SQL = SQL & "  ) "
        SQL = SQL & "  Begin "
        SQL = SQL & "Insert into ServerExport ("
        SQL = SQL & " TableName"
        SQL = SQL & ",TableType"
        SQL = SQL & ",actv_Fg"
        SQL = SQL & " )values ("
        SQL = SQL & AQ(s_TableName)
        SQL = SQL & "," & AQ(s_TableType)
        SQL = SQL & "," & s_ActvFg
        SQL = SQL & ")"
        SQL = SQL & "  End "
        gCnnMst.Execute SQL

End Sub

Private Sub Add2Denoms(s_DenimId As Integer, s_Descr As String, s_DenomStruc As String)
    
        SQL = "If Not Exists " & vbCrLf
        SQL = SQL & " ( Select '1' from Denoms "
        SQL = SQL & "  Where denom_id = " & Val(s_DenimId)
        SQL = SQL & "  ) "
        SQL = SQL & "  Begin "
        SQL = SQL & "Insert into Denoms ("
        SQL = SQL & " denom_id"
        SQL = SQL & ",descr"
        SQL = SQL & ",denom_struc"
        SQL = SQL & " )values ("
        SQL = SQL & Val(s_DenimId)
        SQL = SQL & "," & AQ(s_Descr)
        SQL = SQL & "," & AQ(s_DenomStruc)
        SQL = SQL & ")"
        SQL = SQL & "  End "
        gCnnMst.Execute SQL

End Sub


Public Sub CreatePkgDatabase(s_DbName As String, s_DbPath As String)

    SQL = "CREATE DATABASE " & s_DbName & " ON"
    SQL = SQL & " (   NAME = " & AQ(s_DbName & "_Data")
    SQL = SQL & "     ,FILENAME = " & AQ(s_DbPath & s_DbName & "_Data.MDF")
    SQL = SQL & "     ,SIZE = 1"
    SQL = SQL & "     ,FILEGROWTH = 10%)"
    
    SQL = SQL & " LOG ON"
    SQL = SQL & " (   NAME = " & AQ(s_DbName & "_Log")
    SQL = SQL & "     ,FILENAME = " & AQ(s_DbPath & s_DbName & "_Log.LDF")
    SQL = SQL & "     ,SIZE = 1"
    SQL = SQL & "     ,FILEGROWTH = 10%"
    SQL = SQL & " )"
    SQL = SQL & " COLLATE SQL_Latin1_General_CP1_CI_AS"
    
    gCnnMst.Execute SQL
    
End Sub

'   This function is only use for Programmer only
'   Not to be used in production environment

'This Function Disable MenuEnable/Disable process that will control
'Menus on the basis of whether it is SERVER or TERMINAL
'And it also controls Sales Saving
Public Function IsNeaturalUserMode() As Boolean
    
    If InStr(1, LCase(Command$), LCase("UserMode=Neatural"), vbTextCompare) > 0 Then
        IsNeaturalUserMode = True
    Else
        IsNeaturalUserMode = False
    End If

End Function

Public Function MenuEnableDisable() As Boolean

    If IsNeaturalUserMode Then
        MenuEnableDisable = True
    Else
        MenuEnableDisable = (OperaionMode = enServer)
    End If
    
End Function


Private Sub DefaultUsers()

    'Level      Meaning
    ' 1         ADMIN           - Can perform all operations
    ' 2         IMPORT/EXPORT   - Can perform only Import/Export
    ' 3         TRAINING        - TRAINING Mode All operations available but data is deleted while Training user logged out
    ' 0         USER            - Can operate Sales and Reports
    
    SQL = "Delete from UserMast where Uid in (1001,1002,1003)"
    gCnnMst.Execute SQL
    
    SQL = "Insert into UserMast ("
    SQL = SQL & "Uid,Uname,Pwd,Level"
    SQL = SQL & " ) Values ("
    SQL = SQL & " 1001,'ADMIN'," & AQ(ChartoAsc("1001")) & ",1)"
    gCnnMst.Execute SQL

    SQL = "Insert into UserMast ("
    SQL = SQL & "Uid,Uname,Pwd,Level"
    SQL = SQL & " ) Values ("
    SQL = SQL & " 1002,'IMPEX'," & AQ(ChartoAsc("1002")) & ",2)"
    gCnnMst.Execute SQL

    SQL = "Insert into UserMast ("
    SQL = SQL & "Uid,Uname,Pwd,Level"
    SQL = SQL & " ) Values ("
    SQL = SQL & " 1003,'TRAINING'," & AQ(ChartoAsc("1003")) & ",3)"
    gCnnMst.Execute SQL

End Sub

Public Function UserLevel(s_UserId As String) As enUserLevel
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    SQL = "Select level From UserMast "
    SQL = SQL & " Where Uid = " & Val(s_UserId)
    
    OpenAdoRst rs, SQL
    
    If rs.RecordCount > 0 Then
        UserLevel = IfNullThen(rs.Fields("level").Value, 0)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Public Function IsTrainingMode() As Integer
    
    If UserLevel(gUser) = eTraining Then
        IsTrainingMode = 1
    Else
        IsTrainingMode = 0
    End If
    
End Function

Public Sub DeleteTrainingModeData()
        
    Dim arrTables(15) As String
    Dim i As Integer
    
    arrTables(0) = "BillSequence"
    arrTables(1) = "Categories"
    arrTables(2) = "EventMast"
    arrTables(3) = "Invdet"
    arrTables(4) = "Invtrn"
    arrTables(5) = "Items"
    arrTables(6) = "KeybrdItem"
    arrTables(7) = "KeybrdSetup"
    arrTables(8) = "Locations"
    arrTables(9) = "Saldet"
    arrTables(10) = "Saltrn"
    arrTables(11) = "Sizes"
    arrTables(12) = "TerminalConfig"
    arrTables(13) = "TerSaldet"
    arrTables(14) = "TerSaltrn"
    arrTables(15) = "Units"
    'arrTables(16) = "UserMast"
    
    For i = 0 To UBound(arrTables)
        SQL = "Delete from " & arrTables(i)
        SQL = SQL & " Where Trng_fg = 1"
        gCnnMst.Execute SQL
    Next
    
End Sub

