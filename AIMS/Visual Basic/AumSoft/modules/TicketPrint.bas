Attribute VB_Name = "TicketPrint"
Option Explicit

Public rstPrint As ADODB.Recordset

'POS Printer Realted Declaration
'******************************************************************************
Public mpHandle As Long ' Printer handle for Status API

Public mW_Descr As Double
Public mW_Qty As Double
Public mW_RtlPrc As Double
Public mW_DiscAmt As Double
Public mW_SalesAmt As Double
Public mW_PgHf3_1 As Double


' Constant variables holding the control command characters.
Private Const CTRL_LEFT = "w"
Private Const CTRL_CENTER = "x"
Private Const CTRL_RIGHT = "y"

Public Type udtTktHf
    RptHf1 As String * 40
    RptHf2 As String * 40
    
    PgHf1 As String * 40
    PgHf2 As String * 40
    
    PgHf3_1 As String * 20
    PgHf3_2 As String * 20
    
    PgHf4_1 As String * 20
    PgHf4_2 As String * 20
End Type

'Public Type udtTkt
'    Descr1 As String * 16
'    Descr2 As String * 32
'    Qty As String * 4
'    RtlAmt As String * 7
'    DiscAmt As String * 5
'    SalesAmt As String * 9
'End Type

Public Type udtTkt
    Descr1 As String * 15 '18
    Descr2 As String * 15 '18
    Descr3 As String * 15 '18
    Qty As String * 4
    RtlAmt As String * 7
    DiscAmt As String * 5
    SalesAmt As String * 9
End Type

Public mTkt As udtTkt
Public mTktHf As udtTktHf
'
'******************************************************************************

Private Function RAllign(s_String As String) As String
    
    Dim iFirstSpacePos As Integer
    Dim iLastSpacePos As Integer
    
    's_String = Replace(s_String, ".", "")
    
    iFirstSpacePos = InStr(1, s_String, " ")
    iLastSpacePos = InStrRev(s_String, " ")
    
    If iFirstSpacePos > 0 Then
        '3.11
        RAllign = Space$(((iLastSpacePos - iFirstSpacePos) + 1) * 1) & Trim$(s_String)
    Else
        RAllign = Space$(1) & s_String
    End If
    
End Function

Private Function RAllignTab(s_String As String) As String
    
    Dim iFirstSpacePos As Integer
    
    iFirstSpacePos = InStr(1, s_String, " ")
    
    If iFirstSpacePos > 0 Then
        RAllignTab = vbTab & Trim$(s_String)
    Else
        RAllignTab = Trim$(s_String)
    End If
    
End Function

Private Function TicketHeader(ByRef s_rst As ADODB.Recordset) As String

    SQL = ""
    
    With s_rst
        
        'Vallabh ManavUddharak Mandal
        SQL = Space(30) & "JHC btlJWæ""thf bkz¤" & vbCrLf
        
        'Vyas Pulication
        SQL = SQL & Space(40) & "Ôgtm vrçjfuNl" & vbCrLf
        
        SQL = SQL & String(40, "=") & vbCrLf
        SQL = SQL & ";theF - " & .Fields("DtaDat") & vbTab  'Tarikh
        SQL = SQL & "xefex - " & .Fields("TranId") & vbCrLf 'Ticket Number
        SQL = SQL & String(40, "-") & vbCrLf
    End With
    
    
    SQL = SQL & "krJd;" & vbTab & vbTab 'Vigat
    SQL = SQL & "lkd" & Space(2)        'Nang
    SQL = SQL & "rfkb;" & Space(5)      'Kinmat
    SQL = SQL & "mntg" & Space(16)      'Sahay
    SQL = SQL & "hfb"                   'Rakam
    
    SQL = SQL & vbCrLf & String(40, "-")
    
    TicketHeader = SQL
    
End Function


Private Function TicketFooter(ByRef s_rst As ADODB.Recordset) As String

    Dim mTck As udtTkt

    With s_rst
        mTck.Descr1 = "fwj."        'Kul : Total
        mTck.Qty = Format(.Fields("ItemQty").Value, "####")
        mTck.SalesAmt = Format(.Fields("Sal_Prc").Value, "####.00")
    End With
    
    SQL = ""
    'First Line
    SQL = String(40, "-") & vbCrLf
    SQL = SQL & mTck.Descr1
    SQL = SQL & RAllignTab(mTck.Qty)
    SQL = SQL & Space(40) & RAllign(mTck.SalesAmt) & vbCrLf
    SQL = SQL & String(40, "=")
        
    TicketFooter = SQL
    
End Function

Public Sub PrintPreview(s_TranId As String, Optional bPreview As Boolean = False)
    
    SQL = "Exec stpFetchTicket " & AQ(s_TranId)

    Set rstPrint = New ADODB.Recordset
    rstPrint.Open SQL, gCnnMst
       
    If rstPrint.RecordCount > 0 Then
        If bPreview Then
            Select Case gTktFmt
                Case 3
                    Call PreviewData3
                Case 2
                    Call PreviewData2
                Case Else
                    Call PreviewData
            End Select
        Else
            Select Case gTktFmt
                Case 3
                    Call PrintData3
                Case 2
                    Call PrintData2
                Case Else
                    Call PrintData
            End Select
        End If
    End If

End Sub


Private Sub PrintData()
    Dim yPos, ctr As Integer
    
    With Printer
        .ScaleMode = vbTwips
        .FontSize = 11
        .FontName = gGujaratiFontName
        '.NewPage
    End With
    
    'Ticket Image--------------------------------------------------
'    Printer.FontSize = 10
'    Printer.Font = "Control"
'    Printer.Print CTRL_CENTER
'    Printer.PaintPicture Image1.Picture, 10, 10, , , , , ,  vbMergeCopy
    'Ticket Image--------------------------------------------------
    
    'Ticket Header--------------------------------------------------
    TicketHeaderFooter 1, rstPrint
        Printer.FontSize = 14
        Printer.ForeColor = vbRed
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = True
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 900
        Printer.Print Trim$(mTktHf.RptHf1)
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 750
        Printer.Print Trim$(mTktHf.RptHf2)
        
        
        Printer.FontSize = 11
        Printer.ForeColor = vbBlack
        Printer.Print ""

        yPos = Printer.CurrentY
        
        Printer.CurrentX = 0
        Printer.Print mTktHf.PgHf3_1
        
        'Printer.CurrentY = yPos
        Printer.CurrentX = 0 ' mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        Printer.Print mTktHf.PgHf3_2
    


    TicketHeaderFooter 2, rstPrint
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = False
        
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print mTkt.Descr1
        
        Printer.CurrentY = yPos
        Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
        Printer.Print mTkt.Qty
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
        Printer.Print mTkt.RtlAmt
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - Printer.TextWidth(mTkt.DiscAmt))
        Printer.Print mTkt.DiscAmt
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
        Printer.Print mTkt.SalesAmt
        
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
    'Ticket Header--------------------------------------------------
    
    'Ticket Detail--------------------------------------------------
    While Not rstPrint.EOF
        TicketDetail rstPrint
                
        Select Case Val(rstPrint.Fields("Grp").Value)
            'Detail Section
            Case 1
            
                Printer.ScaleMode = vbTwips
                Printer.FontSize = 11
                Printer.FontName = gGujaratiFontName
                Printer.FontBold = False
                
                'First Line
                yPos = Printer.CurrentY
                Printer.CurrentX = 0
                Printer.Print mTkt.Descr1
                        
                Printer.CurrentY = yPos
                Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
                Printer.Print mTkt.Qty
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
                Printer.Print mTkt.RtlAmt
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - Printer.TextWidth(mTkt.DiscAmt))
                Printer.Print mTkt.DiscAmt
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
                Printer.Print mTkt.SalesAmt
            
                'Second Line
                If Trim(mTkt.Descr2) <> "" Then
                    Printer.CurrentX = 0
                    Printer.Print Trim(mTkt.Descr2)
                End If
                
                'Third Line
                If Trim(mTkt.Descr3) <> "" Then
                    Printer.CurrentX = 0
                    Printer.Print Trim(mTkt.Descr3)
                End If
                
            'Total Section
            Case 2
            
                Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
            
                Printer.ScaleMode = vbTwips
                Printer.FontSize = 12
                Printer.FontName = gGujaratiFontName
                Printer.FontBold = True
                
                yPos = Printer.CurrentY
                Printer.CurrentX = 0
                Printer.Print "  fwj."            'Kul Total
            
                Printer.CurrentY = yPos
                Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
                Printer.Print mTkt.Qty

                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
                Printer.Print mTkt.SalesAmt
            
        End Select
    
        rstPrint.MoveNext
    Wend
    
    Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
    'Ticket Detail--------------------------------------------------
    
    'Ticket Footer--------------------------------------------------
    TicketHeaderFooter 3, rstPrint
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = False
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print mTktHf.PgHf3_1
        
        Printer.CurrentY = yPos
        Printer.CurrentX = mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        Printer.Print mTktHf.PgHf3_2
    
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print mTktHf.PgHf4_1
    
        Printer.FontSize = 11
        Printer.ForeColor = vbRed
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = True
        
        Printer.CurrentX = 900
        Printer.Print mTktHf.PgHf4_2
        
    'Ticket Footer--------------------------------------------------
    
    Printer.CurrentX = 0
    
    Printer.EndDoc
    
End Sub


Public Sub TicketDetail(ByRef s_rst As ADODB.Recordset)

    With s_rst
        mTkt.Descr1 = "*" & .Fields("ItemName").Value
        'mTkt.Descr2 = Space(1) + Mid$(.Fields("ItemName").Value, 17)
        mTkt.Descr2 = Mid$(.Fields("ItemName").Value, 15, 15)
        mTkt.Descr3 = Mid$(.Fields("ItemName").Value, 30, 15)
        
        mTkt.Qty = Format(.Fields("ItemQty").Value, "####")
        mTkt.RtlAmt = Format(.Fields("Rtl_Price").Value, "####0")
        mTkt.DiscAmt = Format(.Fields("Disc_Amt").Value, "##0")
        mTkt.SalesAmt = Format(.Fields("Bill_Amt").Value, "####0")
    End With

End Sub

Public Sub TicketHeaderFooter(iHeaderMode As Integer, ByRef s_rst As ADODB.Recordset)
    
On Error GoTo errhndl
    
    Select Case iHeaderMode
        Case 1      'Report Header
            With s_rst
                'mTktHf.RptHf1 = "JHC btlJutWæ""thf bkz¤"              'Aum Vallabh ManavUddharak Mandal Aum
                mTktHf.RptHf1 = "Î >l-v{Î >l mrbr;"                   'Dan Pradan Samiti
                mTktHf.RptHf2 = "JÕjC>§b, yl>Jj."                     'Vallabhshram Anaval.
                mTktHf.PgHf3_1 = ";t. " & .Fields("DtaDat")           'Tarikh
                mTktHf.PgHf3_2 = "lÝ. " & .Fields("TranId")           'Ticket Number
            End With
        
        Case 2      'Page Header
            With s_rst
                mTkt.Descr1 = "krJd;"               'Vigat
                mTkt.Descr2 = ""
                mTkt.Descr3 = ""
                mTkt.Qty = "lkd"                    'Nang
                mTkt.RtlAmt = "rfkb;"                'Kinmat
                mTkt.DiscAmt = " mntg"              'Sahay
                mTkt.SalesAmt = "  hfb"              'Rakam
            End With
            
        Case 3      'Report Footer
            With s_rst
                .MoveLast
                mTktHf.PgHf3_1 = "awfJKe - " & Format(.Fields("Paid_Amt"), "###0.00")             'Chuknavani  (Customer Payment)
                mTktHf.PgHf3_2 = "vh; - " & Format(.Fields("Change_Amt"), "###0.00")              'Parat              (Change)
                mTktHf.PgHf4_1 = "fwj mntg - " & Format(.Fields("Disc_Amt"), "###0.00")              'Parat              (Change)
                mTktHf.PgHf4_2 = "***sg vhbtðbt***"                                                   'Jay Parmatma
            End With
            
    End Select
    
    Exit Sub
errhndl:
    MsgBox Err.Number & "-" & Err.Description
    
End Sub


Public Function GetPrinter(m_Cmd As CommandButton) As Boolean

    On Error Resume Next
    
    Dim printerObj As Printer
    GetPrinter = False
    
    Const PRINTER_NAME = "EPSON TM-U220* Receipt"
    Dim hndleButton As Long

    For Each printerObj In Printers
       If printerObj.DeviceName Like "*" + PRINTER_NAME + "*" Then
            Set Printer = printerObj
            GetPrinter = True
            Exit For
        End If
    Next

'    mpHandle = BiOpenMonPrinter(TYPE_PRINTER, Printer.DeviceName)
'    If mpHandle < 0 Then
'        GetPrinter = False
'        MsgBox ("Failed to open printer status monitor.")
'    Else
'        hndleButton = m_Cmd.hWnd
'
'        If Not BiSetStatusBackWnd(mpHandle, hndleButton, VarPtr(lpdwStatus)) = SUCCESS Then
'            GetPrinter = False
'            MsgBox ("Failed to set callback function.")
'        End If
'    End If

    'Set Default Widths------------------------------------
    With Printer
        .FontSize = 10
        .ScaleMode = vbTwips
        
'        mW_Descr = .TextWidth(mTkt.Descr1)
'        mW_Qty = .TextWidth(mTkt.Qty)
'        mW_RtlPrc = .TextWidth(mTkt.RtlAmt)
'        mW_DiscAmt = .TextWidth(mTkt.DiscAmt)
'        mW_SalesAmt = .TextWidth(mTkt.SalesAmt)
'        mW_PgHf3_1 = .TextWidth(mTktHf.PgHf3_1)
        
        mW_Descr = 1440
        mW_Qty = 360
        mW_RtlPrc = 630
        mW_DiscAmt = 450
        mW_SalesAmt = 810
        
        mW_PgHf3_1 = 1800

    End With
    
    Set printerObj = Nothing

End Function

Private Sub PreviewData()
    Dim yPos, ctr As Integer
    
    With frmSalesRegister.picPreview
        .ScaleMode = vbTwips
        .FontSize = 11
        .FontName = gGujaratiFontName
    End With
    
    
    'Ticket Header--------------------------------------------------
    TicketHeaderFooter 1, rstPrint
        frmSalesRegister.picPreview.FontSize = 14
        frmSalesRegister.picPreview.ForeColor = vbRed
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = True
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print Trim(mTktHf.RptHf1)
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 750
        frmSalesRegister.picPreview.Print mTktHf.RptHf2
        
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.ForeColor = vbBlack
        frmSalesRegister.picPreview.Print ""

        yPos = frmSalesRegister.picPreview.CurrentY
        
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_1
        
        'frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 0 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_2
    


    TicketHeaderFooter 2, rstPrint
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = False
        
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTkt.Descr1
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
        frmSalesRegister.picPreview.Print mTkt.Qty
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
        frmSalesRegister.picPreview.Print mTkt.RtlAmt
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - frmSalesRegister.picPreview.TextWidth(mTkt.DiscAmt))
        frmSalesRegister.picPreview.Print mTkt.DiscAmt
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
        frmSalesRegister.picPreview.Print mTkt.SalesAmt
        
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
    'Ticket Header--------------------------------------------------
    
    'Ticket Detail--------------------------------------------------
    While Not rstPrint.EOF
        TicketDetail rstPrint
                
        Select Case Val(rstPrint.Fields("Grp").Value)
            'Detail Section
            Case 1
            
                frmSalesRegister.picPreview.ScaleMode = vbTwips
                frmSalesRegister.picPreview.FontSize = 11
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
                frmSalesRegister.picPreview.FontBold = False
                
                'First Line
                yPos = frmSalesRegister.picPreview.CurrentY
                frmSalesRegister.picPreview.CurrentX = 0
                frmSalesRegister.picPreview.Print mTkt.Descr1
                        
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - frmSalesRegister.picPreview.TextWidth(mTkt.Qty))
                frmSalesRegister.picPreview.Print mTkt.Qty
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - frmSalesRegister.picPreview.TextWidth(mTkt.RtlAmt))
                frmSalesRegister.picPreview.Print mTkt.RtlAmt
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - frmSalesRegister.picPreview.TextWidth(mTkt.DiscAmt))
                frmSalesRegister.picPreview.Print mTkt.DiscAmt
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - frmSalesRegister.picPreview.TextWidth(mTkt.SalesAmt))
                frmSalesRegister.picPreview.Print mTkt.SalesAmt
            
                'Second Line
                If Trim(mTkt.Descr2) <> "" Then
                    frmSalesRegister.picPreview.CurrentX = 0
                    frmSalesRegister.picPreview.Print Trim(mTkt.Descr2)
                End If
                
                'Third Line
                If Trim(mTkt.Descr3) <> "" Then
                    frmSalesRegister.picPreview.CurrentX = 0
                    frmSalesRegister.picPreview.Print Trim(mTkt.Descr3)
                End If
                
            'Total Section
            Case 2
            
                frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
            
                frmSalesRegister.picPreview.ScaleMode = vbTwips
                frmSalesRegister.picPreview.FontSize = 12
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
                frmSalesRegister.picPreview.FontBold = True
                
                yPos = frmSalesRegister.picPreview.CurrentY
                frmSalesRegister.picPreview.CurrentX = 0
                frmSalesRegister.picPreview.Print "  fwj."            'Kul Total
            
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - frmSalesRegister.picPreview.TextWidth(mTkt.Qty))
                frmSalesRegister.picPreview.Print mTkt.Qty

                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - frmSalesRegister.picPreview.TextWidth(mTkt.SalesAmt))
                frmSalesRegister.picPreview.Print mTkt.SalesAmt
            
        End Select
    
        rstPrint.MoveNext
    Wend
    
    frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
    'Ticket Detail--------------------------------------------------
    
    'Ticket Footer--------------------------------------------------
    TicketHeaderFooter 3, rstPrint
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = False
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_1
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2) - frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2))
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_2
    
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTktHf.PgHf4_1
    
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.ForeColor = vbRed
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = True
    
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print mTktHf.PgHf4_2
        
    'Ticket Footer--------------------------------------------------
    
    frmSalesRegister.picPreview.CurrentX = 0
    
    
End Sub

Public Sub DisplayMessage()
    If (status And ASB_PRINT_SUCCESS) = ASB_PRINT_SUCCESS Then
        'MsgBox ("Complete printing.")
        'Call OpenDrawer(mpHandle)
                        
    ElseIf (status And ASB_NO_RESPONSE) = ASB_NO_RESPONSE Then
        MsgBox ("No response")
            
    ElseIf (status And ASB_COVER_OPEN) = ASB_COVER_OPEN Then
        MsgBox ("Cover is open.")
            
    ElseIf (status And ASB_AUTOCUTTER_ERR) = ASB_AUTOCUTTER_ERR Then
        MsgBox ("Autocutter error occurred.")
            
    ElseIf ((status And ASB_PAPER_END_FIRST) = ASB_PAPER_END_FIRST) Or ((status And ASB_PAPER_END_SECOND) = ASB_PAPER_END_SECOND) Then
        MsgBox ("Roll paper end sensor: paper not present.")
                
    End If
End Sub


Private Sub PreviewData2()
    Dim yPos, ctr As Integer
    
    With frmSalesRegister.picPreview
        .ScaleMode = vbTwips
        .FontSize = 11
        .FontName = gGujaratiFontName
    End With
    
    
    'Ticket Header--------------------------------------------------
    TicketHeaderFooter2 1, rstPrint
        frmSalesRegister.picPreview.FontSize = 14
        frmSalesRegister.picPreview.ForeColor = vbRed
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = True
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print Trim(mTktHf.RptHf1)
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 750
        frmSalesRegister.picPreview.Print mTktHf.RptHf2
        
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.ForeColor = vbBlack

        yPos = frmSalesRegister.picPreview.CurrentY
        
        'Set English font for amount figures
        frmSalesRegister.picPreview.FontSize = 9
        frmSalesRegister.picPreview.FontName = "Arial"
                
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_2
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 2800 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_1) - Printer.TextWidth(mTktHf.PgHf3_1))
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_1
    
        'Reset font
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName


    TicketHeaderFooter2 2, rstPrint
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = False
        
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTkt.Descr1
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
        frmSalesRegister.picPreview.Print mTkt.Qty
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
        frmSalesRegister.picPreview.Print mTkt.RtlAmt
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - frmSalesRegister.picPreview.TextWidth(mTkt.DiscAmt))
        frmSalesRegister.picPreview.Print mTkt.DiscAmt
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
        frmSalesRegister.picPreview.Print mTkt.SalesAmt
        
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
    'Ticket Header--------------------------------------------------
    
    'Ticket Detail--------------------------------------------------
    While Not rstPrint.EOF
        TicketDetail2 rstPrint
                
        Select Case Val(rstPrint.Fields("Grp").Value)
            'Detail Section
            Case 1
            
                frmSalesRegister.picPreview.ScaleMode = vbTwips
                frmSalesRegister.picPreview.FontSize = 11
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
                frmSalesRegister.picPreview.FontBold = False
                
                'First Line
                yPos = frmSalesRegister.picPreview.CurrentY
                frmSalesRegister.picPreview.CurrentX = 0
                frmSalesRegister.picPreview.Print mTkt.Descr1 & mTkt.Descr2 & mTkt.Descr3
                        
                'Check for Second Line
                If Trim$(mTkt.Descr2) = "" Then
                    frmSalesRegister.picPreview.CurrentY = yPos
                Else
                    yPos = frmSalesRegister.picPreview.CurrentY
                End If
                
                'Set English font for amount figures
                frmSalesRegister.picPreview.FontSize = 9
                frmSalesRegister.picPreview.FontName = "Arial"
                
                frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - frmSalesRegister.picPreview.TextWidth(mTkt.Qty))
                frmSalesRegister.picPreview.Print mTkt.Qty
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - frmSalesRegister.picPreview.TextWidth(mTkt.RtlAmt))
                frmSalesRegister.picPreview.Print mTkt.RtlAmt
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - frmSalesRegister.picPreview.TextWidth(mTkt.DiscAmt))
                frmSalesRegister.picPreview.Print mTkt.DiscAmt
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - frmSalesRegister.picPreview.TextWidth(mTkt.SalesAmt))
                frmSalesRegister.picPreview.Print mTkt.SalesAmt
            
                'Reset font
                frmSalesRegister.picPreview.FontSize = 11
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
           
                
            'Total Section
            Case 2
            
                frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
            
                frmSalesRegister.picPreview.ScaleMode = vbTwips
                frmSalesRegister.picPreview.FontSize = 12
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
                frmSalesRegister.picPreview.FontBold = True
                
                yPos = frmSalesRegister.picPreview.CurrentY
                frmSalesRegister.picPreview.CurrentX = 0
                frmSalesRegister.picPreview.Print "  fwj."            'Kul Total
            
                'Set English font for amount figures
                frmSalesRegister.picPreview.FontSize = 9
                frmSalesRegister.picPreview.FontName = "Arial"
            
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - frmSalesRegister.picPreview.TextWidth(mTkt.Qty))
                frmSalesRegister.picPreview.Print mTkt.Qty

                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - frmSalesRegister.picPreview.TextWidth(mTkt.SalesAmt))
                frmSalesRegister.picPreview.Print mTkt.SalesAmt
            
                'Reset font
                frmSalesRegister.picPreview.FontSize = 11
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
            
        End Select
    
        rstPrint.MoveNext
    Wend
    
    
    'Ticket Detail--------------------------------------------------
    
    'Ticket Footer--------------------------------------------------
    TicketHeaderFooter2 3, rstPrint
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = False
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print "  awfJKe - "
        
        'Set English font for amount figures
        frmSalesRegister.picPreview.FontSize = 9
        frmSalesRegister.picPreview.FontName = "Arial"
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_1
        
        'Reset font
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 1800 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2) - frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2))
        frmSalesRegister.picPreview.Print "vh; - "
    
        'Set English font for amount figures
        frmSalesRegister.picPreview.FontSize = 9
        frmSalesRegister.picPreview.FontName = "Arial"
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 2300 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2) - frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2))
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_2
        
        'Reset font
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print "  mntg - "
    
        'Set English font for amount figures
        frmSalesRegister.picPreview.FontSize = 9
        frmSalesRegister.picPreview.FontName = "Arial"
    
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print mTktHf.PgHf4_1
    
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
    
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.ForeColor = vbRed
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = True
    
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 100
        frmSalesRegister.picPreview.Print """LgJ>Î"
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 2450
        frmSalesRegister.picPreview.Print "sg vhbtðbt"
        
    'Ticket Footer--------------------------------------------------
    
    frmSalesRegister.picPreview.CurrentX = 0
    
    
End Sub

Private Sub PrintData2()
    Dim yPos, ctr As Integer
    
    With Printer
        .ScaleMode = vbTwips
        .FontSize = 11
        .FontName = gGujaratiFontName
    End With
    
    
    'Ticket Header--------------------------------------------------
    TicketHeaderFooter2 1, rstPrint
        Printer.FontSize = 14
        Printer.ForeColor = vbRed
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = True
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 900
        Printer.Print Trim(mTktHf.RptHf1)
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 750
        Printer.Print mTktHf.RptHf2
        
        Printer.FontSize = 11
        Printer.ForeColor = vbBlack

        yPos = Printer.CurrentY
        
        'Set English font for amount figures
        Printer.FontSize = 9
        Printer.FontName = "Arial"
                
        Printer.CurrentX = 0
        Printer.Print mTktHf.PgHf3_2
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 2800 'mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_1) - Printer.TextWidth(mTktHf.PgHf3_1))
        Printer.Print mTktHf.PgHf3_1
    
        'Reset font
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName


    TicketHeaderFooter2 2, rstPrint
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = False
        
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print mTkt.Descr1
        
        Printer.CurrentY = yPos
        Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
        Printer.Print mTkt.Qty
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
        Printer.Print mTkt.RtlAmt
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - Printer.TextWidth(mTkt.DiscAmt))
        Printer.Print mTkt.DiscAmt
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
        Printer.Print mTkt.SalesAmt
        
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
    'Ticket Header--------------------------------------------------
    
    'Ticket Detail--------------------------------------------------
    While Not rstPrint.EOF
        TicketDetail2 rstPrint
                
        Select Case Val(rstPrint.Fields("Grp").Value)
            'Detail Section
            Case 1
            
                Printer.ScaleMode = vbTwips
                Printer.FontSize = 11
                Printer.FontName = gGujaratiFontName
                Printer.FontBold = False
                
                'First Line
                yPos = Printer.CurrentY
                Printer.CurrentX = 0
                Printer.Print mTkt.Descr1 & mTkt.Descr2 & mTkt.Descr3
                        
                'Check for Second Line
                If Trim$(mTkt.Descr2) = "" Then
                    Printer.CurrentY = yPos
                Else
                    yPos = Printer.CurrentY
                End If
                
                'Set English font for amount figures
                Printer.FontSize = 9
                Printer.FontName = "Arial"
                
                Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
                Printer.Print mTkt.Qty
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
                Printer.Print mTkt.RtlAmt
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - Printer.TextWidth(mTkt.DiscAmt))
                Printer.Print mTkt.DiscAmt
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
                Printer.Print mTkt.SalesAmt
            
                'Reset font
                Printer.FontSize = 11
                Printer.FontName = gGujaratiFontName
           
                
            'Total Section
            Case 2
            
                Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
            
                Printer.ScaleMode = vbTwips
                Printer.FontSize = 12
                Printer.FontName = gGujaratiFontName
                Printer.FontBold = True
                
                yPos = Printer.CurrentY
                Printer.CurrentX = 0
                Printer.Print "  fwj."            'Kul Total
            
                'Set English font for amount figures
                Printer.FontSize = 9
                Printer.FontName = "Arial"
            
                Printer.CurrentY = yPos
                Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
                Printer.Print mTkt.Qty

                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
                Printer.Print mTkt.SalesAmt
            
                'Reset font
                Printer.FontSize = 11
                Printer.FontName = gGujaratiFontName
            
        End Select
    
        rstPrint.MoveNext
    Wend
    
    
    'Ticket Detail--------------------------------------------------
    
    'Ticket Footer--------------------------------------------------
    TicketHeaderFooter2 3, rstPrint
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = False
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print "  awfJKe - "
        
        'Set English font for amount figures
        Printer.FontSize = 9
        Printer.FontName = "Arial"
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 900
        Printer.Print mTktHf.PgHf3_1
        
        'Reset font
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 1800 'mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        Printer.Print "vh; - "
    
        'Set English font for amount figures
        Printer.FontSize = 9
        Printer.FontName = "Arial"
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 2300 'mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        Printer.Print mTktHf.PgHf3_2
        
        'Reset font
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print "  mntg - "
    
        'Set English font for amount figures
        Printer.FontSize = 9
        Printer.FontName = "Arial"
    
        Printer.CurrentY = yPos
        Printer.CurrentX = 900
        Printer.Print mTktHf.PgHf4_1
    
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
    
        Printer.FontSize = 11
        Printer.ForeColor = vbRed
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = True
    
        yPos = Printer.CurrentY
        Printer.CurrentX = 100
        Printer.Print """LgJ>Î"
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 2450
        Printer.Print "sg vhbtðbt"
        
    'Ticket Footer--------------------------------------------------
    
    Printer.CurrentX = 0
    
    Printer.EndDoc
    
End Sub

Public Sub TicketHeaderFooter2(iHeaderMode As Integer, ByRef s_rst As ADODB.Recordset)
    
On Error GoTo errhndl
    
    Select Case iHeaderMode
        Case 1      'Report Header
            With s_rst
                mTktHf.RptHf1 = "Î >l-v{Î >l mrbr;"                     'Dan Pradan Samiti
                mTktHf.RptHf2 = "JÕjC>§b, yl>Jj."                       'Vallabhshram Anaval.
                mTktHf.PgHf3_1 = .Fields("DtaDat")                      'Tarikh
                mTktHf.PgHf3_2 = .Fields("TranId")                      'Ticket Number
            End With
        
        Case 2      'Page Header
            With s_rst
                mTkt.Descr1 = "krJd;"               'Vigat
                mTkt.Descr2 = ""
                mTkt.Descr3 = ""
                mTkt.Qty = "lkd"                    'Nang
                mTkt.RtlAmt = "rfkb;"                'Kinmat
                mTkt.DiscAmt = " mntg"              'Sahay
                mTkt.SalesAmt = "  hfb"              'Rakam
            End With
            
        Case 3      'Report Footer
            With s_rst
                .MoveLast
                mTktHf.PgHf3_1 = Format(.Fields("Paid_Amt"), "###0.00")             'Chuknavani  (Customer Payment)
                mTktHf.PgHf3_2 = Format(.Fields("Change_Amt"), "###0.00")           'Parat              (Change)
                mTktHf.PgHf4_1 = Format(.Fields("Disc_Amt"), "###0.00")
               
                mTktHf.PgHf4_2 = "" '"""LgJ>Î sg vhbtðbt"                                                   'Jay Parmatma
            End With
            
    End Select
    
    Exit Sub
errhndl:
    MsgBox Err.Number & "-" & Err.Description
    
End Sub

Public Sub TicketDetail2(ByRef s_rst As ADODB.Recordset)

    With s_rst
        mTkt.Descr1 = "*" & .Fields("ItemName").Value
        'mTkt.Descr2 = Space(1) + Mid$(.Fields("ItemName").Value, 17)
        mTkt.Descr2 = Mid$(.Fields("ItemName").Value, 15, 15)
        mTkt.Descr3 = Mid$(.Fields("ItemName").Value, 30, 15)
        
        mTkt.Qty = Format(.Fields("ItemQty").Value, "####")
        mTkt.RtlAmt = Format(.Fields("Rtl_Price").Value, "####0.00")
        mTkt.DiscAmt = Format(.Fields("Disc_Amt").Value, "##0")
        mTkt.SalesAmt = Format(.Fields("Bill_Amt").Value, "####0.00")
    End With

End Sub

Private Sub PreviewData3()
    Dim yPos, ctr As Integer
    
    With frmSalesRegister.picPreview
        .ScaleMode = vbTwips
        .FontSize = 11
        .FontName = gGujaratiFontName
    End With
    
    
    'Ticket Header--------------------------------------------------
    TicketHeaderFooter2 1, rstPrint
        frmSalesRegister.picPreview.FontSize = 14
        frmSalesRegister.picPreview.ForeColor = vbRed
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = True
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print Trim(mTktHf.RptHf1)
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 750
        frmSalesRegister.picPreview.Print mTktHf.RptHf2
        
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.ForeColor = vbBlack

        yPos = frmSalesRegister.picPreview.CurrentY
                
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_2
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 2550 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_1) - Printer.TextWidth(mTktHf.PgHf3_1))
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_1
    
       

    TicketHeaderFooter2 2, rstPrint
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = False
        
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print mTkt.Descr1
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
        frmSalesRegister.picPreview.Print mTkt.Qty
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
        frmSalesRegister.picPreview.Print mTkt.RtlAmt
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - frmSalesRegister.picPreview.TextWidth(mTkt.DiscAmt))
        frmSalesRegister.picPreview.Print mTkt.DiscAmt
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
        frmSalesRegister.picPreview.Print mTkt.SalesAmt
        
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
    'Ticket Header--------------------------------------------------
    
    'Ticket Detail--------------------------------------------------
    While Not rstPrint.EOF
        TicketDetail2 rstPrint
                
        Select Case Val(rstPrint.Fields("Grp").Value)
            'Detail Section
            Case 1
            
                frmSalesRegister.picPreview.ScaleMode = vbTwips
                frmSalesRegister.picPreview.FontSize = 11
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
                frmSalesRegister.picPreview.FontBold = False
                
                'First Line
                yPos = frmSalesRegister.picPreview.CurrentY
                frmSalesRegister.picPreview.CurrentX = 0
                frmSalesRegister.picPreview.Print mTkt.Descr1 & mTkt.Descr2 & mTkt.Descr3
                        
                'Check for Second Line
                If Trim$(mTkt.Descr2) = "" Then
                    frmSalesRegister.picPreview.CurrentY = yPos
                Else
                    yPos = frmSalesRegister.picPreview.CurrentY
                End If
                
                frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - frmSalesRegister.picPreview.TextWidth(mTkt.Qty))
                frmSalesRegister.picPreview.Print mTkt.Qty
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - frmSalesRegister.picPreview.TextWidth(mTkt.RtlAmt))
                frmSalesRegister.picPreview.Print mTkt.RtlAmt
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - frmSalesRegister.picPreview.TextWidth(mTkt.DiscAmt))
                frmSalesRegister.picPreview.Print mTkt.DiscAmt
                
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - frmSalesRegister.picPreview.TextWidth(mTkt.SalesAmt))
                frmSalesRegister.picPreview.Print mTkt.SalesAmt
            
                
            'Total Section
            Case 2
            
                frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
            
                frmSalesRegister.picPreview.ScaleMode = vbTwips
                frmSalesRegister.picPreview.FontSize = 12
                frmSalesRegister.picPreview.FontName = gGujaratiFontName
                frmSalesRegister.picPreview.FontBold = True
                
                yPos = frmSalesRegister.picPreview.CurrentY
                frmSalesRegister.picPreview.CurrentX = 0
                frmSalesRegister.picPreview.Print "  fwj."            'Kul Total
            
                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = mW_Descr + (mW_Qty - frmSalesRegister.picPreview.TextWidth(mTkt.Qty))
                frmSalesRegister.picPreview.Print mTkt.Qty

                frmSalesRegister.picPreview.CurrentY = yPos
                frmSalesRegister.picPreview.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - frmSalesRegister.picPreview.TextWidth(mTkt.SalesAmt))
                frmSalesRegister.picPreview.Print mTkt.SalesAmt
            
        End Select
    
        rstPrint.MoveNext
    Wend
    
    
    'Ticket Detail--------------------------------------------------
    
    'Ticket Footer--------------------------------------------------
    TicketHeaderFooter2 3, rstPrint
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = False
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print "  awfJKe - "
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_1
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 1800 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2) - frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2))
        frmSalesRegister.picPreview.Print "vh; - "
    
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 2300 'mW_PgHf3_1 + (frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2) - frmSalesRegister.picPreview.TextWidth(mTktHf.PgHf3_2))
        frmSalesRegister.picPreview.Print mTktHf.PgHf3_2
        
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 0
        frmSalesRegister.picPreview.Print "  mntg - "
    
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 900
        frmSalesRegister.picPreview.Print mTktHf.PgHf4_1
    
        frmSalesRegister.picPreview.Line (10, frmSalesRegister.picPreview.CurrentY)-(3600, frmSalesRegister.picPreview.CurrentY)
    
        frmSalesRegister.picPreview.FontSize = 11
        frmSalesRegister.picPreview.ForeColor = vbRed
        frmSalesRegister.picPreview.FontName = gGujaratiFontName
        frmSalesRegister.picPreview.FontBold = True
    
        yPos = frmSalesRegister.picPreview.CurrentY
        frmSalesRegister.picPreview.CurrentX = 100
        frmSalesRegister.picPreview.Print """LgJ>Î"
        
        frmSalesRegister.picPreview.CurrentY = yPos
        frmSalesRegister.picPreview.CurrentX = 2450
        frmSalesRegister.picPreview.Print "sg vhbtðbt"
        
    'Ticket Footer--------------------------------------------------
    
    frmSalesRegister.picPreview.CurrentX = 0
    
    
End Sub

Private Sub PrintData3()
    Dim yPos, ctr As Integer
    
    With Printer
        .ScaleMode = vbTwips
        .FontSize = 11
        .FontName = gGujaratiFontName
    End With
    
    
    'Ticket Header--------------------------------------------------
    TicketHeaderFooter2 1, rstPrint
        Printer.FontSize = 14
        Printer.ForeColor = vbRed
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = True
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 900
        Printer.Print Trim(mTktHf.RptHf1)
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 750
        Printer.Print mTktHf.RptHf2
        
        Printer.FontSize = 11
        Printer.ForeColor = vbBlack

        yPos = Printer.CurrentY
        
        Printer.CurrentX = 0
        Printer.Print mTktHf.PgHf3_2
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 2550 'mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_1) - Printer.TextWidth(mTktHf.PgHf3_1))
        Printer.Print mTktHf.PgHf3_1
    

    TicketHeaderFooter2 2, rstPrint
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = False
        
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print mTkt.Descr1
        
        Printer.CurrentY = yPos
        Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
        Printer.Print mTkt.Qty
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
        Printer.Print mTkt.RtlAmt
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - Printer.TextWidth(mTkt.DiscAmt))
        Printer.Print mTkt.DiscAmt
        
        Printer.CurrentY = yPos
        Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
        Printer.Print mTkt.SalesAmt
        
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
    'Ticket Header--------------------------------------------------
    
    'Ticket Detail--------------------------------------------------
    While Not rstPrint.EOF
        TicketDetail2 rstPrint
                
        Select Case Val(rstPrint.Fields("Grp").Value)
            'Detail Section
            Case 1
            
                Printer.ScaleMode = vbTwips
                Printer.FontSize = 11
                Printer.FontName = gGujaratiFontName
                Printer.FontBold = False
                
                'First Line
                yPos = Printer.CurrentY
                Printer.CurrentX = 0
                Printer.Print mTkt.Descr1 & mTkt.Descr2 & mTkt.Descr3
                        
                'Check for Second Line
                If Trim$(mTkt.Descr2) = "" Then
                    Printer.CurrentY = yPos
                Else
                    yPos = Printer.CurrentY
                End If
                
                Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
                Printer.Print mTkt.Qty
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty) + (mW_RtlPrc - Printer.TextWidth(mTkt.RtlAmt))
                Printer.Print mTkt.RtlAmt
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc) + (mW_DiscAmt - Printer.TextWidth(mTkt.DiscAmt))
                Printer.Print mTkt.DiscAmt
                
                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
                Printer.Print mTkt.SalesAmt
            
                
            'Total Section
            Case 2
            
                Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
            
                Printer.ScaleMode = vbTwips
                Printer.FontSize = 12
                Printer.FontName = gGujaratiFontName
                Printer.FontBold = True
                
                yPos = Printer.CurrentY
                Printer.CurrentX = 0
                Printer.Print "  fwj."            'Kul Total
            
                Printer.CurrentY = yPos
                Printer.CurrentX = mW_Descr + (mW_Qty - Printer.TextWidth(mTkt.Qty))
                Printer.Print mTkt.Qty

                Printer.CurrentY = yPos
                Printer.CurrentX = (mW_Descr + mW_Qty + mW_RtlPrc + mW_DiscAmt) + (mW_SalesAmt - Printer.TextWidth(mTkt.SalesAmt))
                Printer.Print mTkt.SalesAmt
            
        End Select
    
        rstPrint.MoveNext
    Wend
    
    
    'Ticket Detail--------------------------------------------------
    
    'Ticket Footer--------------------------------------------------
    TicketHeaderFooter2 3, rstPrint
        Printer.FontSize = 11
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = False
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print "  awfJKe - "
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 900
        Printer.Print mTktHf.PgHf3_1
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 1800 'mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        Printer.Print "vh; - "
    
        Printer.CurrentY = yPos
        Printer.CurrentX = 2300 'mW_PgHf3_1 + (Printer.TextWidth(mTktHf.PgHf3_2) - Printer.TextWidth(mTktHf.PgHf3_2))
        Printer.Print mTktHf.PgHf3_2
        
        yPos = Printer.CurrentY
        Printer.CurrentX = 0
        Printer.Print "  mntg - "
    
        Printer.CurrentY = yPos
        Printer.CurrentX = 900
        Printer.Print mTktHf.PgHf4_1
    
        Printer.Line (10, Printer.CurrentY)-(3600, Printer.CurrentY)
    
        Printer.FontSize = 11
        Printer.ForeColor = vbRed
        Printer.FontName = gGujaratiFontName
        Printer.FontBold = True
    
        yPos = Printer.CurrentY
        Printer.CurrentX = 100
        Printer.Print """LgJ>Î"
        
        Printer.CurrentY = yPos
        Printer.CurrentX = 2450
        Printer.Print "sg vhbtðbt"
        
    'Ticket Footer--------------------------------------------------
    
    Printer.CurrentX = 0
    
    Printer.EndDoc
    
End Sub

