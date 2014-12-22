VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTestInterface 
   Caption         =   "Test Interface"
   ClientHeight    =   8460
   ClientLeft      =   2550
   ClientTop       =   2070
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11850
   Begin VB.CommandButton Command8 
      Caption         =   "First page"
      Height          =   540
      Left            =   240
      TabIndex        =   20
      Top             =   4815
      Width           =   1725
   End
   Begin VB.TextBox txtBarcodeSize 
      Height          =   480
      Left            =   315
      TabIndex        =   19
      Text            =   "3"
      Top             =   2265
      Width           =   465
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Web Browser"
      Height          =   510
      Left            =   9000
      TabIndex        =   18
      Top             =   4815
      Width           =   1770
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Bar Code"
      Height          =   420
      Left            =   6705
      TabIndex        =   17
      Top             =   4860
      Width           =   1920
   End
   Begin VB.CommandButton Command5 
      Caption         =   "End Of Day"
      Height          =   465
      Left            =   3690
      TabIndex        =   16
      Top             =   4785
      Width           =   2445
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send Debit String"
      Height          =   375
      Left            =   6645
      TabIndex        =   15
      Top             =   4395
      Width           =   2025
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PIN Pad Handshake"
      Height          =   375
      Left            =   8955
      TabIndex        =   14
      Top             =   4305
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PIN Pad GetTField"
      Height          =   375
      Left            =   8955
      TabIndex        =   12
      Top             =   3780
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Receipt"
      Height          =   675
      Left            =   8940
      TabIndex        =   11
      Top             =   2985
      Width           =   2025
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   195
      Top             =   1545
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin Threed.SSCommand cmdPrinter 
      Height          =   645
      Left            =   6630
      TabIndex        =   10
      Top             =   3675
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1138
      _Version        =   196609
      Caption         =   "CommPrinter"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   195
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3705
      Width           =   3045
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   675
      Left            =   150
      TabIndex        =   8
      Top             =   4080
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1191
      _Version        =   196609
      Caption         =   "Send to Ticketmaster"
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   390
      Left            =   3630
      TabIndex        =   7
      Top             =   3720
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   688
      _Version        =   196609
      Caption         =   "PIN Pad Debit"
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   600
      Left            =   6585
      TabIndex        =   6
      Top             =   3015
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1058
      _Version        =   196609
      Caption         =   "Printer"
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   630
      Left            =   150
      TabIndex        =   5
      Top             =   2985
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   1111
      _Version        =   196609
      Caption         =   "Login TicketMaster"
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   285
      Top             =   885
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "29.11.45.21"
      RemotePort      =   1130
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   615
      Left            =   3660
      TabIndex        =   4
      Top             =   2985
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   "Swipe Card"
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   4305
      TabIndex        =   3
      Top             =   2370
      Width           =   5955
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   4305
      TabIndex        =   1
      Top             =   1785
      Width           =   5970
   End
   Begin Threed.SSCommand SSCommand9 
      Height          =   390
      Left            =   3660
      TabIndex        =   13
      Top             =   4260
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   688
      _Version        =   196609
      Caption         =   "PIN Pad Credit"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Approval Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   1230
      TabIndex        =   2
      Top             =   2430
      Width           =   2940
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Response Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   1230
      TabIndex        =   0
      Top             =   1845
      Width           =   2940
   End
End
Attribute VB_Name = "frmTestInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCommand As String
Dim CardReturn As String
Dim BarcodeSize As String
Function ConstructBarCodeNY(TotalTicketNumber As Integer, ByRef thisTicketType As Integer) As String
    Dim i                       As Integer
    Dim j                       As Integer
    Dim K                       As Integer
    Dim strBarcodeString        As String
'        JComputerID = "A"
'        j = Len(Trim(JournalID))
'        strBarcodeString = "1"
'        If UCase(Trim(JComputerID)) = "A" Then strBarcodeString = "1"
'        If UCase(Trim(JComputerID)) = "B" Then strBarcodeString = "2"
'        If UCase(Trim(JComputerID)) = "C" Then strBarcodeString = "3"
'        If UCase(Trim(JComputerID)) = "D" Then strBarcodeString = "4"
'        If UCase(Trim(JComputerID)) = "E" Then strBarcodeString = "5"
'        For i = 1 To 9 - j
'             strBarcodeString = strBarcodeString & "0"
'        Next i
'        strBarcodeString = strBarcodeString & Trim(JournalID)
'        strBarcodeString = Left(JournalID, 2) & "Z" & Right(JournalID, 3)
        Select Case thisTicketType
        Case 1
            strBarcodeString = "81" & Right(SellerCode, 2) & Format(Right(CStr(SessionID), 6), "00000#")
        Case 3
            strBarcodeString = "82" & Right(SellerCode, 2) & Format(Right(CStr(SessionID), 6), "00000#")
        Case 16
             strBarcodeString = "83" & Right(SellerCode, 2) & Format(Right(CStr(SessionID), 6), "00000#")
        End Select
        
'        strBarcodeString = "88" & Right(SellerCode, 2) & Format(Right(CStr(SessionID), 6), "00000#")
'        For i = 1 To 8
'            strBarcodeString = strBarcodeString & Trim(CStr(TotalTicketNumber))
'        Next i
        If TotalTicketNumber >= 99 Then TotalTicketNumber = 0
           TicketNumStr = CStr(TotalTicketNumber)
        If (Len(TicketNumStr) < 2) Then TicketNumStr = "0" & TicketNumStr
        strBarcodeString = strBarcodeString & TicketNumStr
   
        ConstructBarCodeNY = strBarcodeString



End Function



Sub PrintTicketsTest(i As Integer, K As Integer)
            Dim rsHeaders As Recordset
            Dim thisType As Integer
            Dim TicketTypes(5, 1)
            ''bDebug = True
            ShowID = "45"
            ''SETUP
            thisType = 1
            TicketTypes(1, i) = 1
            TicketTypes(2, i) = 1
            TicketTypes(3, i) = 5
            TicketTypes(4, i) = 1
            TicketTypes(5, i) = "ADULT"
            EventCodeID = 1
            EventCode = "NYAU08"
            TicketCount = 1
            SellerCode = "ATMT01"
            JournalID = "1234567"
            JComputerID = 1
            MOP = "VISA"
            EPDATE = Now()
            BarCodeType = "0"
            
            
            Set rsHeaders = db.OpenRecordset("Select * from TICKET_HEADERS where ShowID=" & ShowID)

            '''''''
            strPrint = Chr(2) & Chr(11) & Chr(2) & Chr(11) & Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) ''    {clear printer buffer}
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(8) & Chr(1) & Chr(2) & Chr(8) & Chr(5)  ''    //init barcode Vertical
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(16) & "2" & ConstructBarCodeNY(K, thisType)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(16) & Chr(8) & rsHeaders("Location") '' & " " & UCase(Location_name)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(2) & Format(TicketTypes(3, i), "$#.00")
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(11) & UCase(rsHeaders("Header1")) 'UCase(ln(1))"
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(34) & Format(TicketTypes(3, i), "$#.00")
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(2) & UCase(EventCode)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(11) & UCase(rsHeaders("Header2"))  'UCase(ln(2))"
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(34) & UCase(EventCode)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(19) & Chr(2) & UCase(TicketTypes(5, i))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(19) & Chr(11) & UCase(rsHeaders("Header3")) 'UCase(ln(3))"
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(19) & Chr(34) & UCase(TicketTypes(5, i))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(2) & " TIX#" & CStr(K)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(11) & UCase(rsHeaders("Header4"))  ''UCase(ln(4))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(34) & " TIX#" & CStr(K)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(2) & UCase(Left(SellerCode, 6))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(11) & UCase(EPDATE)
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(34) & UCase(Left(SellerCode, 6))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(1) & UCase(Trim(EventCode))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(10) & Format(Now(), "m/d/yyyy") & " " & MOP
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(34) & UCase(Trim(EventCode))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            ''MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(23) & Chr(11) & rsHeaders("Header5")  ''UCase(ln(5))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(23) & Chr(34) & "A" & Trim(JournalID) ''& UCase(Trim(JComputerID) & Trim(JournalID))
            If bDebug Then MsgBox strPrint
            MSComm1.Output = strPrint
            If bDebug Then MsgBox "PRINT"
            MSComm1.Output = Chr(2) & "pn"

End Sub


Function ConstructBarCode0(TotalTicketNumber As Integer, JournalID As String, JComputerID As String) As String
    Dim i                       As Integer
    Dim j                       As Integer
    Dim K                       As Integer
    Dim strBarcodeString        As String
'        JComputerID = "A"
        j = Len(Trim(JournalID))
        strBarcodeString = "1"
        If UCase(Trim(JComputerID)) = "A" Then strBarcodeString = "1"
        If UCase(Trim(JComputerID)) = "B" Then strBarcodeString = "2"
        If UCase(Trim(JComputerID)) = "C" Then strBarcodeString = "3"
        If UCase(Trim(JComputerID)) = "D" Then strBarcodeString = "4"
        If UCase(Trim(JComputerID)) = "E" Then strBarcodeString = "5"
        For i = 1 To 9 - j
             strBarcodeString = strBarcodeString & "0"
        Next i
        strBarcodeString = strBarcodeString & Trim(JournalID)
        If TotalTicketNumber >= 99 Then TotalTicketNumber = 0
        'TicketNumStr = CStr(TotalTicketNumber + 1)
        TicketNumStr = CStr(TotalTicketNumber)
        If (Len(TicketNumStr) < 2) Then TicketNumStr = "0" & TicketNumStr
        strBarcodeString = strBarcodeString & TicketNumStr
        ConstructBarCode0 = strBarcodeString

End Function
Sub PrintReceipt()
        Dim strPrint As String
            ADDRESS1 = "3560 CARNATION CIRCLE"
            ADDRESS2 = "SEAL BEACH"
            ADDRESS3 = "CA"
            bankmessage = "BANK MESSAGE"
            totalDollar = "10.00"
            CustomAccttype = "CUSTTYPE"
            ret_ref_num = "1111"
            ApproveCode = "1"
            TerminalID = "ABCD"
            CardNumber = "1234567890"
            CardExpire = "06/05"
            MSComm1.CommPort = 4
            MSComm1.Settings = "9600,N,8,1"
            ' Open the port.
            MSComm1.PortOpen = True
            strPrint = Chr(2) & Chr(11) & Chr(2) & Chr(11) & Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) ''    {clear printer buffer}
            l = Len(strPrint)
            For i = 1 To l
                ' kb[0] := InitPrintString[i];
                ' kb[1] := char(0);
                 MSComm1.Output = Mid(strPrint, i, 1) & Chr(0)
            Next i
            
            strPrint = Chr(13) & Chr(10) & Chr(2) & Chr(15) & Chr(13) & Chr(13) & Chr(2) & Chr(23) & Chr(15) & Chr(20) & Chr(1)
'            MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(11) & Chr(2) & Chr(11) & "@BDEFGIKO@**Q***"
'            MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(13) & Chr(10) & Chr(0)
'            MsgBox strPrint
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(15) & Chr(13) & Chr(10) & Chr(2) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(0)
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(16) & Chr(1) & ADDRESS1
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(1) & ADDRESS2
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(1) & ADDRESS3
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(1) & "ACCT#:"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(7) & CardNumber
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(1) & " PURCHASE"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(23) & Chr(1) & bankmessage
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(28) & " DEBIT RECORD"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(28) & " HOST DATE"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(40) & " TIME"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(28) & Format(Now(), "m/d/yyyy")
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(38) & Format(Now(), "hh:mm")
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(29 + 1) & " AMT"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(34 + 6) & " $" & totalDollar
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(17) & Chr(47) & CustomAccttype
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(47) & " EXPIRY:"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(18) & Chr(55) & CardExpire
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(19) & Chr(50) & " RRN"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(20) & Chr(47) & ret_ref_num
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(44 + 3) & " AUTH#:"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(21) & Chr(51 + 3) & ApproveCode
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(22) & Chr(44 + 3) & " TERM#:"
            MSComm1.Output = strPrint
            strPrint = Chr(2) & Chr(23) & Chr(44 + 3) & TerminalID
            MSComm1.Output = strPrint
            MSComm1.Output = Chr(2) & "pn"
            MSComm1.PortOpen = False

End Sub


Function GenerateCommand(K As Integer) As String
'Dim strData As String
'Dim strData2 As String
'Dim Discount(10) As String
'Dim DiscountName(10) As String
'Dim i As Integer
'Discount(0) = "A"
'Discount(1) = "B"
'Discount(2) = "C"
'Discount(3) = "D"
'Discount(4) = "E"
'
'DiscountName(0) = "ADULT"
'DiscountName(1) = "SENIOR"
'DiscountName(2) = "STUDENT"
'DiscountName(3) = ""
'If bTest Then
'     EventType = "TESTATM"
'Else
'     EventType = "CANADA BLOOMS"
'End If
'
'Select Case K
'    Case 0
'     strData = "LOGI " & Trim(SellerCode) & "    " & Trim(Password) & "                                  TORONTO                      @#end#@"
'   Case 1
'       strData = "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:TORONTO" & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"
'       'Winsock1.SendData "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:TORONTO" & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"
'    Case "SEVT"
'       Winsock1.GetData strData
'       strCommand = "EVNT"
'       i = 1
'        Do Until InStr(i, strData, "TEST") = 0
'            i = InStr(i + 1, strData, Chr(10))
'        Loop
'        If i = 1 Then i = InStr(1, strData, Chr(10))
'        EventID = Mid(strData, i - 54, 10)
'       If bDebug Then MsgBox "EVENTID: " & EventID
'       'EventID = Mid(strData, 61, 10)
'       Winsock1.SendData "EVNT " & SellerCode & "    " & Password & "    " & Chr(10) & "City: TORONTO" & Chr(10) & "Event:" & EventID & "A" & Chr(10) & "@#end#@"
'    Case "EVNT" 'SHOWS for a CATEGORY return
'        Winsock1.GetData strData
'        'Today Show
'        If bDebug Then MsgBox "SHOW INFO" & Chr(10) & strData
'        EventCode = Mid(strData, InStr(1, strData, "EVENTCODE:") + 10, (InStr(1, strData, "LOCATION") - 1) - (InStr(1, strData, "EVENTCODE:") + 10))
'        'If Len(EventCode) < 6 Then EventCode = EventCode & "01"
'        If Len(EventCode) < 16 Then
'            EventCode = Trim(EventCode) & Space(16 - Len(Trim(EventCode)))
'        Else
'
'        End If
'        If bDebug Then MsgBox "EventCode:" & EventCode
'        strCommand = "OENT" 'ORDER TICKETS
'        'strCommand = "XXX" 'ORDER TICKETS
'        strData = "OENT " & SellerCode & "    " & Password & "    " & Chr(10)
'        For i = 0 To UBound(TicketTypes, 2)
'            If TotalTicketSold(i) > 0 Then
'               'strData = strData & "Item:" & EventID & "A          " & EventCode & "          `   P1    " & Discount(i) & CStr(TotalTicketSold(i)) & "        " & DiscountName(i) & " / Price1" & EventCode & "                           " & Chr(10)
'                strData = strData & "Item:" & EventID & "A          " & EventCode & "   P1    " & Discount(i) & CStr(TotalTicketSold(i)) & "        " & DiscountName(i) & " / Price1" & EventCode & "                           " & Chr(10)
'            End If
'        Next i
'        strData = strData & "@#end#@"
'        If bDebug Then MsgBox "ORDER TICKETS" & Chr(10) & strData
'        'If bDebug Then TextRespond.Text = strData
'        'If bDebug Then TextRespond.Visible = True
'
'        Winsock1.SendData strData
'    Case "OENT" 'ORDER TICKETS return
'        Dim TotalPrice As Double
'        strCommand = "CENT" 'CONFIRM TICKETS
''        If CardType = "020" Or CardType = "010" Then
''            MOP = "AMEX"
''        ElseIf CardType = "101" Then
''            MOP = "MC"
''        ElseIf CardType = "201" Then
''            MOP = "VISA"
''        Else
''            MOP = "OTHER"
''        End If
'        If bDebug Then MsgBox "Credit Card left 4 digits: " & Left(CardNumber, 4)
'        If Left(CardNumber, 2) = "34" Or Left(CardNumber, 2) = "37" Then
'          MOP = "AMX"
'        ElseIf Left(CardNumber, 1) = "5" Then
'            MOP = "MC"
'        ElseIf Left(CardNumber, 1) = "4" Then
'            MOP = "VISA"
'        ElseIf Left(CardNumber, 3) = "300" Or Left(CardNumber, 3) = "303" Or Left(CardNumber, 3) = "302" Or Left(CardNumber, 3) = "303" Or Left(CardNumber, 3) = "304" Or Left(CardNumber, 3) = "305" Or Left(CardNumber, 2) = "36" Or Left(CardNumber, 2) = "38" Then
'            MOP = "DSC"
'        ElseIf Left(CardNumber, 4) = "2131" Or Left(CardNumber, 4) = "1800" Then
'            MOP = "JCB"
'        ElseIf Left(CardNumber, 6) = "628181" Then
'            MOP = "SEARS"
'        Else
'            MOP = "OTHER"
'        End If
'
'
'        'if bDebug then MsgBox MOP
'        Winsock1.GetData strData2
'        'if bDebug then textrespond.Text = strData2
'        If bDebug Then MsgBox "ORDER TICKETS response" & Chr(10) & strData2
'        JournalID = Trim(Mid(strData2, InStr(1, strData2, "JournalID") + 10, 10))
'        'if bDebug then MsgBox JournalID
'        strData = "CENT " & SellerCode & "    " & Password & "    " & Chr(10)
'        For i = 0 To UBound(TicketTypes, 2)
'            If TotalTicketSold(i) > 0 Then
'                TotalPrice = TotalPrice + TotalTicketSold(i) * TicketTypes(3, i)
'                'strData = strData & "Item:" & EventID & "A          TEST01             P1     A2          $2.00     2        0.00      0.00      ADULT / PRICE TEST01                    " & Chr(10) & "ComputerID:A" & Chr(10) & "JournalID:933429" & Chr(10) & "MOP:VISA" & Chr(10) & "Name:VISA/GM." & Chr(10) & "CCNum:4551210041000373" & Chr(10) & "Expiry:06/05" & Chr(10) & "GTOTAL:12.00" & Chr(10) & "TOTAL:12.00" & Chr(10) & "GSTTOTAL:0.00" & Chr(10) & "PSTTOTAL:0.00" & Chr(10) & "@#end#@"
'                 strData = strData & "Item:" & EventID & "A         " & EventCode & "    P1    " & Discount(i) & CStr(TotalTicketSold(i)) & "          " & Format(TicketTypes(3, i), "$#.00") & Space(10 - Len(Format(TicketTypes(3, i), "$#.00"))) & Format(TotalTicketSold(i) * TicketTypes(3, i), "#.00") & Space(10 - Len(Format(TotalTicketSold(i) * TicketTypes(3, i), "#.00"))) & "0.00      0.00      " & DiscountName(i) & " / Price1 TEST01                           " & Chr(10)
'                'strData = strData & "Item:" & EventID & "A         " & EventCode & "           P1    " & Discount(i) & CStr(TotalTicketSold(i)) & "        " & Format(TicketTypes(3, i), "$#.00") & Space(10 - Len(Format(TicketTypes(3, i), "$#.00"))) & Format(TotalTicketSold(i) * TicketTypes(3, i), "#.00") & Space(10 - Len(Format(TotalTicketSold(i) * TicketTypes(3, i), "#.00"))) & "0.00      0.00      " & DiscountName(i) & " / Price1 TEST01                           " & Chr(10)
'           End If
'        Next i
'        strData = strData & "ComputerID:A" & Chr(10)
'        strData = strData & "JournalID:" & JournalID & Chr(10)
'        strData = strData & "MOP:" & MOP & Chr(10)
'        strData = strData & "Name:" & MOP & Chr(10)
'        strData = strData & "CCNum:" & CardNumber & Chr(10)
'        strData = strData & "Expiry:" & Left(CardExpire, 2) & "/" & Right(CardExpire, 2) & Chr(10)
'        strData = strData & "GTOTAL:" & Format(TotalPrice, "#.00") & Chr(10)
'        strData = strData & "TOTAL:" & Format(TotalPrice, "#.00") & Chr(10)
'        strData = strData & "GSTTOTAL:0.00" & Chr(10)
'        strData = strData & "PSTTOTAL:0.00" & Chr(10)
'        strData = strData & "@#end#@"
'        Response = 0
'        If bDebug Then
'            Text1.Visible = True
'            Text1.Text = strData
'        End If
'        If bDebug Then Response = MsgBox("CONFIRM TICKETS" & Chr(10) & strData, vbOKCancel)
'        'If bDebug Then TextRespond.Text = strData
'        'If bDebug Then TextRespond.Visible = True
'        If Response = vbCancel Then
'            lblMessage.Caption = "Your Tickets are now PRINTING..."
'            Refresh
'            If bDebug Then MsgBox "Printing"
'                TicketCount = "1111"
'                EventCodeID = "11"
'                PrintTickets
'                Timer1.Interval = 3000
'                Timer1.Enabled = True
'        Else
'            Winsock1.SendData strData
'        End If
'    Case "CENT"  'CONFIRM TICKETS return
'        'if bDebug then MsgBox "CENT"
'        Winsock1.GetData strData
'        If bDebug Then MsgBox "CONFIRM TICKETS response" & Chr(10) & strData
'        'If bDebug Then MsgBox Mid(strData, InStr(1, strData, "AUTHCODE:") + 9, 1)
'        If InStr(1, strData, "CONFIRMEVENTOK") > 0 Then
'                Winsock1.Close
'                'If bDebug Then MsgBox Mid(strData, InStr(1, strData, "AUTHCODE:") + 9, 1)
'                If Trim(Mid(strData, InStr(1, strData, "AUTHCODE:") + 9, 1)) = "1" Or Trim(Mid(strData, InStr(1, strData, "AUTHCODE:") + 9, 1)) = "3" Then
'                    lblMessage.Caption = "Your Tickets are now PRINTING..."
'                    Refresh
'                    If bDebug Then MsgBox "Printing"
'                    i = InStr(InStr(1, strData, "TICKETCOUNT") + 12, strData, Chr(10)) - (InStr(1, strData, "TICKETCOUNT") + 12)
'                    TicketCount = Mid(strData, InStr(1, strData, "TICKETCOUNT") + 12, i)
'                    i = InStr(InStr(1, strData, "EVENTCODEID") + 12, strData, Chr(10)) - (InStr(1, strData, "EVENTCODEID") + 12)
'                    EventCodeID = Mid(strData, InStr(1, strData, "EVENTCODEID") + 12, i)
'                    PrintTickets
'                    Timer1.Interval = 3000
'                    Timer1.Enabled = True
'                ElseIf Trim(Mid(strData, InStr(1, strData, "AUTHCODE:") + 9, 1)) = "2" Then
'                    lblMessage.Caption = "Credit card not accepted..."
'                    Timer1.Interval = 3000
'                    Timer1.Enabled = True
'                ElseIf Trim(Mid(strData, InStr(1, strData, "AUTHCODE:") + 9, 1)) = "4" Then
'                    lblMessage.Caption = "Processing problem - please try again..."
'                    Timer1.Interval = 3000
'                    Timer1.Enabled = True
'
'                End If
'        Else
'            lblMessage.Caption = "Processing problem - please try again..."
'            Timer2.Interval = 3000
'            Timer2.Enabled = True
'        End If
'
'End Select
'Text1(9).Text = strData
End Function


Private Sub cmdPrinter_Click()
Dim strPrint As String
BarcodeSize = txtBarcodeSize.Text
MSComm1.CommPort = 4
    'MSComm1.PortOpen = True
   ' 9600 baud, no parity, 8 data, and 1 stop bit.
   'MsgBox "Port Set"
MSComm1.Settings = "9600,N,8,1"
   'MSComm1.OutBufferSize = 800
   ' Tell the control to read entire buffer when Input
   ' is used.
   ' MsgBox "Settings Set"
  'MSComm1.InputLen = 0
   ' Open the port.
MSComm1.PortOpen = True
   'MsgBox "Port Open"
Call PrintTicketsTest(1, 1)
   
'strPrint = "<RL><RC360,10><F3><HW1,1>GHOSTWRITER WORLD"
'strPrint = strPrint & "<RC380,76><F6><HW1,1><BS26,44>ALL<F2>"
'strPrint = strPrint & "<F6><BS26,44>THREE<F2>  <F6><BS26,44>PARKS"
'strPrint = strPrint & "<RC348,130><F6><HW1,1><BS42,44>PASSPORT"
'strPrint = strPrint & "<RC324,240><RL><F6><HW2,2>6"
'strPrint = strPrint & "<RC210,240><HW1,1>DAY"
'strPrint = strPrint & "<RC230,290><F3><HW1,1>ADMIT ONE"
'strPrint = strPrint & "<RC230,320><F6><HW1,1>GUEST"
'strPrint = strPrint & "<F1><RC230,370><F1><HW1,1>VERY SMALL PRINT"
'strPrint = strPrint & "<RC24,530><LT2><BX340,50>"
'strPrint = strPrint & "<RC25,528><LT2><VX338>"
'strPrint = strPrint & "<RC216,550><HW1,1><F2>DAY  1"
'strPrint = strPrint & "<RC24,580><LT2><BX340,50>"
'strPrint = strPrint & "<RC216,600>DAY  2"
'strPrint = strPrint & "<RC24,630><LT2><BX340,50>"
'strPrint = strPrint & "<RC216,650>DAY  3"
'strPrint = strPrint & "<RC24,680><LT2><BX340,50>"
'strPrint = strPrint & "<RC216,700>DAY  4"
'
'
'   ' Send the attention command to the modem.
'   MsgBox CStr(Len(strPrint))
'   MSComm1.Output = strPrint
'strPrint = "<RC24,730><LT2><BX340,50>"
'strPrint = strPrint & "<RC25,780><LT2><VX338>"
'strPrint = strPrint & "<RC216,760>DAY  5"
'strPrint = strPrint & "<RC340,400><RL><F6><BS36,44><HW1,1>DAY GUEST"
'strPrint = strPrint & "<RC260,450><F3><HW1,1>$112.00"
'strPrint = strPrint & "<RC240,482><F3>PLUS TAX"
'strPrint = strPrint & "<RC280,1010><F3><HW1,1>12345678"
'strPrint = strPrint & "<RC60,990><NL10><X2>*01000407*"
'strPrint = strPrint & "<RC360,820><F9><HW1,1>VALID ONLY ON DATE STAMPED"
'strPrint = strPrint & "NONTRANSFERABLE NONREFUNDABLE"
'strPrint = strPrint & "<RC280,870><F3><HW1,1>01000407"
'strPrint = strPrint & "<RC20,1079><RR><F3><HW1,1>GHOSTWRITER WORLD"
 'strPrint = Text1(9).Text
 'strPrint = "<RC50,20><F3><HW1,1>$12.00<RC100,20><F3><HW1,1>Price1<RC150,20><F3><HW2,1>ADULT<RC250,20><F3><HW1,1>ATM01<RC300,20><F3><HW1,1>04CAS<RC10,160><F3><HW2,1>**CANADIAN INTL AUTO SHOW**<RC80,230><F3><HW1,1>Feb 13-22, 2004<RC120,200><F3><HW1,1>VALID FOR ONE DAY ONLY<RC160,190><F3><HW2,1>CANADIAN INTL AUTO SHOW<RC230,210><F7><HW1,1>FEB 18, 2004 02:30 PM<RC260,220><F7><HW1,1>02/18/2004 VISA<RC300,160><F3><HW1,1>VALID FOR FEB 13-22, 2004"

 'strPrint = Chr(2) & Chr(16) & "2" & "1234567890" & "pn"
 'MsgBox strPrint
 
MSComm1.Output = strPrint
'strPrint = "<RC50,700><F3><HW1,1>$12.00<RC100,700><F3><HW1,1>Price1<RC150,700><F3><HW2,1>ADULT<RC250,700><F3><HW1,1>ATM01<RC300,700><F3><HW1,1>04CAS<RL><RC280,870><F3><HW1,1>123456789012<RC60,990><NL10><X2>*123456789012*"
' strPrint = strPrint & "<p>"
' MsgBox CStr(Len(strPrint))

'MSComm1.Output = strPrint
  
   
   'MSComm1.Output = strPrint2
   'MSComm1.Output = "<p>"
   ' Ensure that
   ' the modem responds with "OK".
   ' Wait for data to come back to the serial port.
'   Do
'      DoEvents
'   buffer$ = buffer$ & MSComm1.Input
'   'Loop Until InStr(Buffer$, "OK" & vbCrLf)
'   Loop Until Len(buffer$) > 0
   
   ' Read the "OK" response data in the serial port.
   ' Close the serial port.
   
   MSComm1.PortOpen = False
'   Text1(9).Text = CStr(Asc(buffer$))

End Sub

Private Sub Combo1_Click()
Dim strData As String
Select Case Combo1.ListIndex
Case 0
    'strData = "LOGI " & Trim(SellerCode) & "    " & Trim(Password) & "                                  TORONTO                      @#end#@"
    strData = "LOGI " & Trim(SellerCode) & "    " & Trim(Password) & "                                  TORONTO                      @#end#@"
    strCommand = "LOGI"
Case 1
    Dim EventType As String
    EventType = "TEST"
    strData = "SEVT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:TORONTO" & Chr(10) & "Category:" & EventType & Chr(10) & "@#end#@"
Case 2
    strCommand = "EVNT"
    strData = "EVNT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:TORONTO   " & Chr(10) & "Event:873       A" & Chr(10) & "@#end#@"
Case 3
    'strData = "OENT TEST05    TEST05    " & Chr(10) & "Item:399       A          TEST01             P1    A1         Adult / Price1 TEST01                           " & Chr(10) & "@#end#@"
    EventID = "873       "
    EventCode = "TEST"
    If Len(EventCode) < 19 Then
        EventCode = Trim(EventCode) & Space(19 - Len(Trim(EventCode)))
    Else
        EventCode = Left(EventCode, 19)
    End If
    DiscountCode = "A"
    TotalTicketsSold = 1
    strData = "OENT " & SellerCode & "    " & Password & "    " & Chr(10)
'    For i = 0 To UBound(TicketTypes, 2)
'        If TotalTicketSold(i) > 0 Then
            strData = strData & "Item:" & EventID & "A          " & EventCode & "P1    A1        ADULT / Price1" & EventCode & "                           " & Chr(10)
 '       End If
 '   Next i
    strData = strData & "@#end#@"
'    strData = "OENT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10)
'    strData = strData & "Item:" & EventID & "A          " & EventCode & "           P1    " & DiscountCode & CStr(TotalTicketsSold) & "        Adult / Price1 " & EventCode & "                         " & Chr(10)
'    strData = strData & "@#end#@"
Case 4
    strData = "CENT " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "Item:873       A          TESTATM            P1    A1          $1.00     1        0.00      0.00      ADULT / PRICE TEST                    " & Chr(10) & "ComputerID:A" & Chr(10) & "JournalID:149486" & Chr(10) & "MOP:VISA" & Chr(10) & "Name:GOODMAN/HS" & Chr(10) & "CCNum:4352375031248299" & Chr(10) & "Expiry:06/08" & Chr(10) & "GTOTAL:1.00" & Chr(10) & "TOTAL:1.00" & Chr(10) & "GSTTOTAL:0.00" & Chr(10) & "PSTTOTAL:0.00" & Chr(10) & "@#end#@"
Case 5
    strData = "CATE " & Trim(SellerCode) & "    " & Trim(Password) & "    " & Chr(10) & "City:TORONTO" & Chr(10) & "@#end#@"

End Select
Text1(9).Text = strData
End Sub


Private Sub Command1_Click()
PrintReceipt
End Sub


Private Sub Command2_Click()
Text1(8).Text = Dsintx1.GetTfield(Text1(9).Text)

End Sub

Private Sub Command3_Click()
'Dsintx1.SetTfield "PinPadPort", 1
Dsintx1.SetTfield "hostname", "ol3d.eigendev.com"
Dsintx1.SetTfield "hostport", "18001"
Dsintx1.SetTfield "DoSSL", "Y"

Dsintx1.InitTransaction
Dsintx1.SetTfield "Trans_Code", 95
'Dsintx1.SetTfield "Term_ID", "TELUSD01"
xreturn = Dsintx1.ProcessTransaction
Text1(9).Text = xreturn
Text1(8).Text = Dsintx1.GetTfield("Response_Code")
Dsintx1.ConfirmTransaction
End Sub


Private Sub Command4_Click()
Dim CardReturn As String
CardReturn = ">513 0 2 3 0 17 2 A12 502 AC0 3 B241 Y2 ID1 Y3 PPF118 B50" & Chr(10) & "D1" & Chr(10) & "G27406871" & Chr(10) & "Q72499919" & Chr(10) & "b0A228E4926A30EFB" & Chr(10) & "d123456     3" & Chr(10) & "h0010010010190" & Chr(10) & "q;4519028522603304=23091204440300000201?" & Chr(10) & "U0" & Chr(10) & "3 PPH23 Z.37TELUSD01        00 2 T239 ;4519028522603304=23091204440300000201?2 TC2 012 TI8 TCMTCC165 api258 000000005 app258 000000004 ds2516 00000000000000004 dt2512 0410130954384 mt252 014 rf258 000001158 src_ex2516 00000000000000004 st252 000 9 3 ECP1 y2 HN13 206.87.131.272 HP4 65309 ISALLEGRO1 Y3 ISD1 Y2 MC1 Y3 PPP1 13 TID8 TCMTCC162 TO2 30"

Winsock1.RemoteHost = IP
'Winsock1.RemoteHost = "29.11.29.72"  'Debit Server
'Winsock1.RemoteHost = "29.11.45.21"  'Ticket Server
Winsock1.RemotePort = 1130
'strCommand = "LOGDEBIT"
Winsock1.Connect


End Sub

Private Sub Command5_Click()
Dim EOFReturn As String
Dsintx1.InitTransaction
Dsintx1.SetTfield "Trans_Code", "62"
EOFReturn = Dsintx1.PrepareTransaction
Text1(8).Text = EOFReturn

End Sub

Private Sub Command7_Click()
frmWebBrowser.Show 1
''frmANAuthorization.Show 1
End Sub

Private Sub Command8_Click()
frmFirstPage.Show
End Sub

Private Sub Form_Load()
Dim x As String
Dim Y As Integer
Dim i As Integer
Dim i2 As Integer
Dim ShowReference As String
Dim strData As String
    ShowReference = "BAYSTATE BRIDAL"
    'x = "Multiple Dates      BOSTON HOME SHOW                        961       A   BAYSIDE EXPO CENTER                     " & Chr(10) & "Multiple Dates      CHILDRENS EXPO                          965       A   BAYSIDE EXPO CENTER                     " & Chr(10) & "@#end#@"
    x = "Multiple Dates      BAYSTATE BRIDAL                         959       A   BAYSIDE EXPO CENTER                     " & Chr(10) & "@#end#@"
    i = InStr(1, x, ShowReference) + 40
    EventID = Mid(x, i, 10)
'    'x = "CENT ADMT01    ATMT01    " & Chr(10) & "Item:399       A          TEST01             P1    A1         $12.00    12        0.00      0.00      ADULT / PRICE TEST01                    " & Chr(10) & "ComputerID:A" & Chr(10) & "JournalID:149486" & Chr(10) & "MOP:VISA" & Chr(10) & "Name:VISA/GM." & Chr(10) & "CCNum:4551210041000373" & Chr(10) & "Expiry:06/05" & Chr(10) & "GTOTAL:12.00" & Chr(10) & "TOTAL:12.00" & Chr(10) & "GSTTOTAL:0.00" & Chr(10) & "PSTTOTAL:0.00" & Chr(10) & "@#end#@"
'    'x = "Item: 1    X $12.00   ticket(s) for $12.00    include GST 0.00    PST 0.00   " & Chr(10) & "TOTAL: 12.00     " & Chr(10) & "" & Chr(10) & "GTOTAL: 12.00     " & Chr(10) & "GSTTOTAL: 0.00      " & Chr(10) & "PSTTOTAL: 0.00      " & Chr(10) & "ComputerID: A                   " & Chr(10) & "JournalID: 149486              " & Chr(10) & "@#end#@"
'      x = "CONFIRMEVENTOK" & Chr(10) & "AUTHRESULT:TIMEOUT" & Chr(10) & "AUTHCODE:3" & Chr(10) & "BARCODETYPE:2" & Chr(10) & "TICKETCOUNT:1259" & Chr(10) & "EVENTCODEID:94" & Chr(10) & "TIME:02/13/2004 08:37:02:030" & Chr(10) & "@#end#@"
'    'x = "Mar  2 2004  2:35PM AUTOPART                                52        A   MTCC                                    "
'    x = "Multiple Dates      TEST EVENT ONLY                         746       A   MTCC                                    " & Chr(10) & "Multiple Dates      2004 AUTO SHOW - GA                     704       A   METRO CONVENTION CENTRE                 " & Chr(10) & "Multiple Dates      AUTO SHOW - DIS                         721       A   METRO CONVENTION CENTRE                 " & Chr(10) & "@#end#@"
'    'x = "Multiple Dates      2004 AUTO SHOW - GA                     705       A   METRO CONVENTION CENTRE                 " & Chr(10) & "Multiple Dates      AUTO SHOW - DIS                         721       A   METRO CONVENTION CENTRE                 " & Chr(10) & "Multiple Dates      TEST EVENT ONLY                         747       A   MTCC                                    " & Chr(10) & "@#end#@"
'    y = 1
'    Do Until InStr(y, x, "TEST") = 0
'        y = InStr(y + 1, x, Chr(10))
'    Loop
'    EventID = Trim(Mid(x, y - 54, 10))
'    x = "Multiple Dates      2004 CANADA BLOOMS                      755       A   METRO CONVENTION CTR                    " & Chr(10) & "@#end#@"
'    x = "%B3728 944600 04019^GOODMAN/HS                ^0405020151497?"
'x = "CONFIRMEVENTOK" & Chr(10) & "AUTHRESULT:TIMEOUT" & Chr(10) & "AUTHCODE:3" & Chr(10) & "BARCODETYPE:2" & Chr(10) & "TICKETCOUNT:1174" & Chr(10) & "EVENTCODEID:100" & Chr(10) & "TIME:02/03/2004 09:31:13:503" & Chr(10) & "@#end#@"
'strData = x
''CardName = Trim(Mid(strData, InStr(1, strData, "^") + 1, InStr((InStr(1, strData, "^") + 1), strData, "^") - InStr(1, strData, "^") - 1))
''i = 1
'' Do Until InStr(i, strData, "TEST") = 0
''     i = InStr(i + 1, strData, Chr(10))
'' Loop
'' If i = 1 Then i = InStr(i + 1, strData, Chr(10))
'' EventID = Mid(strData, i - 54, 10)
'y = InStr(InStr(1, strData, "TICKETCOUNT") + 12, strData, Chr(10)) - (InStr(1, strData, "TICKETCOUNT") + 12)
'TicketCount = Mid(strData, InStr(1, strData, "TICKETCOUNT") + 12, y)
'i = InStr(InStr(1, strData, "EVENTCODEID") + 12, strData, Chr(10)) - (InStr(1, strData, "EVENTCODEID") + 12)
'EventCodeID = Mid(strData, InStr(1, strData, "EVENTCODEID") + 12, i)

'strData = "Multiple Dates      A ONE OF A KIND - SPRING                762       A   NATIONAL TRADE CENTRE" & Chr(10) & "Multiple Dates      B ONE OF A KIND - SPRING                767       A   NATIONAL TRADE CENTRE" & Chr(10) & "@#end#@"
' strData = "Multiple Dates      TORONTO TEST EVENT                      400       A   MTCC NORTH & SOUTH                      " & Chr(10) & "@#end#@"
 'strData = "Multiple Dates      A ONE OF A KIND - SPRING                762       A   NATIONAL TRADE CENTRE                   " & Chr(10) & "Multiple Dates      B ONE OF A KIND - SPRING                767       A   NATIONAL TRADE CENTRE                   " & Chr(10) & "Mar 31 2004  5:00PM CANCER SOCIETY - A                      795       A   NATIONAL TRADE CENTRE                   " & Chr(10) & "Mar 31 2004  5:00PM CANCER SOCIETY - B                      796       A   NATIONAL TRADE CENTRE                   " & Chr(10) & "@#end#@"

'       i = 1
'        Do Until InStr(i, strData, Chr(10)) = 0
'            i = InStr(i, strData, "TORONTO TEST EVENT")
'        Loop
'        EventID = Mid(strData, i + 40, 10)
'strData = ">467 0 2 3 0 13 8 Amount_12 108 ReadCard1 Y7 Term_ID8 TCMTCC1510 Track2_Acc31 ;372894460004019=0405020151497?10 Trans_Code2 0112 apicref_ex258 0000000012 appcref_ex258 0000000013 datetime_ex2513 10404021535429 dest_ex2516 000000000000000010 mtype_ex252 018 ref_ex258 000069338 src_ex2516 000000000000000012 subtype_ex252 000 9 7 ENCRYPT1 y8 HOSTNAME13 206.87.131.278 HOSTPORT4 65309 ISALLEGRO1 Y11 ISDECOUPLED1 Y7 MCREDIT1 Y10 PINPADPORT1 17 TERM_ID8 TCMTCC157 TIMEOUT2 30"
'i = InStr(1, strData, ";") + 1
'i2 = InStr(1, strData, "=")
'CardNumber = Mid(strData, i, i2 - i)
'CardExpire = Mid(strData, i2 + 1, 4)
'CardType = Mid(strData, i2 + 5, 3)


'Text1(9).Text = ConstructBarCode0(1, "990032", "A")
Combo1.AddItem "LOGI"
Combo1.AddItem "SEVT"
Combo1.AddItem "EVNT"
Combo1.AddItem "OENT"
Combo1.AddItem "CENT"
Combo1.AddItem "CATE"

End Sub


Private Sub SSCommand1_Click()
Dim xreturn As Boolean
Dsintx1.SetTfield "hostname", Text1(1).Text
Dsintx1.SetTfield "hostport", Text1(2).Text
Dsintx1.SetTfield "DoSSL", "Y"
Dsintx1.InitTransaction
Dsintx1.SetTfield "Trans_Code", Text1(0).Text
Dsintx1.SetTfield "Track2_Acc", Text1(4).Text
Dsintx1.SetTfield "ExpiryDate", Text1(5).Text
Dsintx1.SetTfield "Amount_1", Text1(6).Text
Dsintx1.SetTfield "Term_ID", "EIGENMTSX1"
Dsintx1.SetTfield "Eigen_ID", Text1(7).Text
xreturn = Dsintx1.ProcessTransaction
Text1(8).Text = Dsintx1.GetTfield("Response_Code")
Text1(9).Text = Dsintx1.GetTfield("Approval_Cd")
Dsintx1.ConfirmTransaction

End Sub


Private Sub SSCommand2_Click()
Dim xreturn As Boolean
Dim CardReturn As String
On Error GoTo XXX
'''Dsintx1.SetTfield "PinPadPort", 1
'''Dsintx1.InitTransaction
'''Dsintx1.SetTfield "Trans_Code", "01"
'''Dsintx1.SetTfield "Track2_Acc", ""
''''Dsintx1.SetTfield "Term_ID", "TCMTCC15"
'''Dsintx1.SetTfield "Amount_1", "10"
'''Dsintx1.SetTfield "ReadCard", "y"
''''xreturn = Dsintx1.ProcessTransaction
''''MsgBox xreturn
'''CardReturn = Dsintx1.PrepareTransaction()
'''MsgBox CardReturn
''''Text1(8).Text = CardReturn
''''If xreturn Then
''''If Len(CardReturn) > 0 Then
''''    i = InStr(1, CardReturn, ";") + 1
''''    'MsgBox i
''''    i2 = InStr(1, CardReturn, "=")
''''     'MsgBox i2
''''    CardNumber = Mid(CardReturn, i, i2 - i)
''''    MsgBox CardNumber
''''    CardExpire = Mid(CardReturn, i2 + 1, 4)
''''     MsgBox CardExpire
''''   CardType = Mid(CardReturn, i2 + 5, 3)
''''    MsgBox CardType
''''    Track1 = Dsintx1.GetTfield("Track1")
''''    MsgBox Track1
''''    Text1(9).Text = Track1
''''    CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
''''    MsgBox CardName
''''    Dsintx1.InitTransaction
''''    Dsintx1.ReleaseConn
''''    'AccountNumber = Dsintx1.GetTfield("Track2_Acc")
''''    'If Len(AccountNumber) > 0 Then
'''''        CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
'''''        CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
'''''        CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
'''''        Dsintx1.SetTfield "PinPadPort", 1
'''''        Dsintx1.InitTransaction
'''''        Dsintx1.SetTfield "Track2_Acc", AccountNumber
'''''        xreturn = Dsintx1.ProcessTransaction
''''        'MsgBox xreturn
'''        Text1(8).Text = Dsintx1.GetTfield("Track1")
'''        Text1(9).Text = Dsintx1.GetTfield("Track2_Acc")Dsintx1.SetTfield "PinPadPort", 1

''''    'End If
''''Else
''''        Text1(8).Text = "No Response"
''''End If
''''Dsintx1.SetTfield "Track2_Acc", AccountNumber
''''Dsintx1.ProcessTransaction
''''xreturn = Dsintx1.ProcessTransaction
''''Text1(8).Text = Dsintx1.GetTfield("Track1")
''''Text1(9).Text = Dsintx1.GetTfield("Track2_Acc")
xreturn = False
Dsintx1.SetTfield "PinPadPort", 1
'Dsintx1.SetTfield "PinPadPort", 5
Dsintx1.SetTfield "NoWriteLog", "Y"
Dsintx1.InitTransaction
Dsintx1.SetTfield "ReadCard", "y"
xreturn = Dsintx1.ProcessTransaction
    Text1(8).Text = Dsintx1.GetTfield("Track2_Acc")
    Text1(9).Text = Dsintx1.GetTfield("Track1")
XXX:

End Sub


Private Sub SSCommand3_Click()
    Winsock1.RemoteHost = IP
    Winsock1.RemotePort = 1130
    Winsock1.Connect
    
    
    
End Sub
Private Sub SSCommand4_Click()
Dim LineNo(8) As Integer
Dim FontSize(8) As Integer
Dim RightSide As Integer

RightSide = 5100
LineNo(1) = 200
LineNo(2) = 500
LineNo(3) = 900
LineNo(4) = 1200
LineNo(5) = 200
LineNo(6) = 1600
LineNo(7) = 1900
LineNo(8) = 2300

FontSize(1) = 14
FontSize(2) = 12
FontSize(3) = 12
FontSize(4) = 14
FontSize(5) = 12
FontSize(6) = 14
FontSize(7) = 12
FontSize(8) = 12
With Printer
    .Copies = 2
    
    .Orientation = 2
    
    .Font = "Ariel Narrow"
    .FontBold = True
    
    .FontSize = FontSize(2)
    .CurrentX = 50
    .CurrentY = LineNo(2)
    Printer.Print "$16.00"
    
    .FontSize = FontSize(3)
    .CurrentX = 50
    .CurrentY = LineNo(3)
    Printer.Print "Price1"
    
    .FontSize = FontSize(4)
    .CurrentX = 50
    .CurrentY = LineNo(4)
    Printer.Print "ADULT"
    
    .FontSize = FontSize(6)
    .CurrentX = 50
    .CurrentY = LineNo(6)
    Printer.Print strComputerName
    
    .FontSize = FontSize(7)
    .CurrentX = 50
    .CurrentY = LineNo(7)
    Printer.Print "04CIAS"
    
    .FontSize = FontSize(1)
    .CurrentX = 1100
    .CurrentY = LineNo(1)
    Printer.Print "** CANADIAN INTL AUTO SHOW **"
    
    .FontSize = FontSize(2)
    .CurrentX = 1500
    .CurrentY = LineNo(2)
    Printer.Print "FEB 13 - 22, 2004"
    
    .FontSize = FontSize(3)
    .CurrentX = 1200
    .CurrentY = LineNo(3)
    Printer.Print "VALID FOR ONE DAY ONLY"
    
    .FontSize = 12
    .CurrentX = 1200
    .CurrentY = LineNo(4)
    Printer.Print "CANADIAN INTL AUTO SHOW"
    'Printer.Print "**2004 AUTO SHOW**"
    
    .FontSize = FontSize(6)
    .CurrentX = 1200
    .CurrentY = LineNo(6)
    Printer.Print "FEB 13, 2004  10:30 AM"
    
    .FontSize = FontSize(7)
    .CurrentX = 1200
    .CurrentY = LineNo(7)
    Printer.Print "02/13/04 MC"
    
    .FontSize = FontSize(8)
    .CurrentX = 1200
    .CurrentY = LineNo(8)
    Printer.Print "VALID FOR FEB 13 - 22, 2004"
    
    .FontSize = FontSize(2)
    .CurrentX = RightSide
    .CurrentY = LineNo(2)
    Printer.Print "$16.00"
    
    .FontSize = FontSize(3)
    .CurrentX = RightSide
    .CurrentY = LineNo(3)
    Printer.Print "Price1"
    
    .FontSize = FontSize(4)
    .CurrentX = RightSide
    .CurrentY = LineNo(4)
    Printer.Print "ADULT"
    
    .FontSize = FontSize(6)
    .CurrentX = RightSide
    .CurrentY = LineNo(6)
    Printer.Print "ATMT10"
    
    .FontSize = FontSize(7)
    .CurrentX = RightSide
    .CurrentY = LineNo(7)
    Printer.Print "04CIAS"
    
    .FontSize = FontSize(8)
    .CurrentX = RightSide
    .CurrentY = LineNo(8)
    Printer.Print "AAAAAAA"
    .EndDoc
End With
End Sub

Private Sub SSCommand5_Click()
Dim xreturn As Boolean
'Dsintx1.SetTfield "PinPadPort", 1
'Dsintx1.InitTransaction
'Dsintx1.SetTfield "ReadCard", "y"
'xreturn = Dsintx1.ProcessTransaction
'If xreturn Then
'    AccountNumber = Dsintx1.GetTfield("Track2_Acc")
    'AccountNumber = ";4506300111103=16071200000000100100001?"  'Tony
    AccountNumber = ";4519028522603304=23091204440300000201?"  'Al
    'If Len(AccountNumber) > 0 Then
    '    MsgBox AccountNumber
    '    CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
     '   CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
     '   CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
     '   CardExpire = Right(CardExpire, 2) & "/" & Left(CardExpire, 2)
     '   MsgBox CardNumber
     '   MsgBox CardExpire
        
'        Dsintx1.SetTfield "PinPadPort", 1
'        Dsintx1.SetTfield "encrypt", "N"
'        Dsintx1.SetTfield "PinPadPort", 1
'        Dsintx1.SetTfield "IsAllegro", "Y"
'        Dsintx1.SetTfield "MCredit", "Y"
'        Dsintx1.SetTfield "IsDecoupled", "Y"
        Dsintx1.InitTransaction
        Dsintx1.SetTfield "Trans_Code", "01"
        Dsintx1.SetTfield "Track2_Acc", AccountNumber
        'Dsintx1.SetTfield "Term_ID", "TCMTCC15"
        Dsintx1.SetTfield "Term_ID", "TCMTCC16"
        'Dsintx1.SetTfield "Card_Num", CardNumber
        'Dsintx1.SetTfield "Exp_Date", CardExpire
        Dsintx1.SetTfield "Amount_1", "50"
        'xreturn = Dsintx1.ProcessTransaction
        'MsgBox xreturn
        CardReturn = Dsintx1.PrepareTransaction()
        Text1(8).Text = CardReturn
        MsgBox CardReturn
        'If Len(Text1(9).Text) = 0 Then
        '    Text1(9).Text = "No Response"
        'End If
        'Winsock1.RemoteHost = "29.11.29.72"  'Debit Server
        Winsock1.RemoteHost = "29.11.45.22"  'Ticket Server
        Winsock1.RemotePort = 1130
        strCommand = "LOGDEBIT"
        Winsock1.Connect
'        strCommand = "DBT"
'        MsgBox "DBIT P" & CardReturn
'        Winsock1.SendData "DBIT P" & CardReturn
        
        'Text1(8).Text = CardReturn
        'Text1(9).Text = "DBITP" & CardReturn
        
        
'        Dsintx1.ConfirmTransaction
        
        'MsgBox xreturn
        'Text1(8).Text = Dsintx1.GetTfield("Response_Code")
        'Text1(9).Text = Dsintx1.GetTfield("ReceiptRefNumber")
'    End If
'End If
'Dsintx1.SetTfield "Track2_Acc", AccountNumber
'Dsintx1.ProcessTransaction
'xreturn = Dsintx1.ProcessTransaction
'Text1(8).Text = Dsintx1.GetTfield("Track1")
'Text1(9).Text = Dsintx1.GetTfield("Track2_Acc")
End Sub

Private Sub SSCommand6_Click()

Winsock1.SendData Text1(9).Text



End Sub

Private Sub SSCommand7_Click()
Dim x As String
x = "Item: 2    X $1.00    ticket(s) for $2.00     include GST 0.00    PST 0.00   " & Chr(10) & "Item: 3    X $1.00    ticket(s) for $3.00     include GST 0.00    PST 0.00   " & Chr(10) & "Item: 1    X $1.00    ticket(s) for $1.00     include GST 0.00    PST 0.00   " & Chr(10) & "TOTAL: 6.00      " & Chr(10) & "GTOTAL: 6.00      " & Chr(10) & "GSTTOTAL: 0.00      " & Chr(10) & "PSTTOTAL: 0.00      " & Chr(10) & "ComputerID: A                   " & Chr(10) & "JournalID: 891968              " & Chr(10) & "@#end#@"

Text1(9).Text = x
End Sub



Private Sub SSCommand9_Click()
Dim xreturn As Boolean
'Dsintx1.SetTfield "PinPadPort", 1
'Dsintx1.InitTransaction
'Dsintx1.SetTfield "ReadCard", "y"
'xreturn = Dsintx1.ProcessTransaction
'If xreturn Then
'    AccountNumber = D4506300111103sintx1.GetTfield("Track2_Acc")
    AccountNumber = ";372894460004019=0405020151497?"
    If Len(AccountNumber) > 0 Then
        MsgBox AccountNumber
        CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
        CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
        CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
        CardExpire = Right(CardExpire, 2) & "/" & Left(CardExpire, 2)
        MsgBox CardNumber
        MsgBox CardExpire
        
        Dsintx1.SetTfield "PinPadPort", 1
        Dsintx1.InitTransaction
        Dsintx1.SetTfield "Trans_Code", "01"
        Dsintx1.SetTfield "Track2_Acc", AccountNumber
        Dsintx1.SetTfield "Term_ID", "TCMTCC15"
        Dsintx1.SetTfield "Card_Num", CardNumber
        Dsintx1.SetTfield "Exp_Date", CardExpire
        Dsintx1.SetTfield "Amount_1", "10000"
        'xreturn = Dsintx1.ProcessTransaction
        'MsgBox xreturn
        Text1(8).Text = Dsintx1.PrepareTransaction()
        If Len(Text1(8).Text) = 0 Then
            Text1(8).Text = "No Response"
        End If
        'MsgBox xreturn
        'Text1(8).Text = Dsintx1.GetTfield("Response_Code")
        'Text1(9).Text = Dsintx1.GetTfield("ReceiptRefNumber")
'    End If
End If
'Dsintx1.SetTfield "Track2_Acc", AccountNumber
'Dsintx1.ProcessTransaction
'xreturn = Dsintx1.ProcessTransaction
'Text1(8).Text = Dsintx1.GetTfield("Track1")
'Text1(9).Text = Dsintx1.GetTfield("Track2_Acc")
End Sub

Private Sub Winsock1_Connect()
'Dim CardReturn As String
'CardReturn = ">513 0 2 3 0 17 2 A12 502 AC0 3 B241 Y2 ID1 Y3 PPF118 B50" & Chr(10) & "D1" & Chr(10) & "G27406871" & Chr(10) & "Q72499919" & Chr(10) & "b0A228E4926A30EFB" & Chr(10) & "d123456     3" & Chr(10) & "h0010010010190" & Chr(10) & "q;4519028522603304=23091204440300000201?" & Chr(10) & "U0" & Chr(10) & "3 PPH23 Z.37TELUSD01        00 2 T239 ;4519028522603304=23091204440300000201?2 TC2 012 TI8 TCMTCC165 api258 000000005 app258 000000004 ds2516 00000000000000004 dt2512 0410130954384 mt252 014 rf258 000001158 src_ex2516 00000000000000004 st252 000 9 3 ECP1 y2 HN13 206.87.131.272 HP4 65309 ISALLEGRO1 Y3 ISD1 Y2 MC1 Y3 PPP1 13 TID8 TCMTCC162 TO2 30"

MsgBox "CONNECT"
SellerCode = "ATMT16"
Password = "ATMT16"
'MsgBox "LOGIN: " & Chr(10) & "LOGI " & SellerCode & "    " & Password & "                                  BOSTON                      @#end#@"
'   strData = "LOGI " & Trim(SellerCode) & "    " & Trim(Password) & "                                  BOSTON                      @#end#@"
''''Winsock1.SendData "LOGI " & SellerCode & "    " & Password & "                                  TORONTO                      @#end#@"
'    strCommand = "LOGI"
''''    MsgBox strData
'    Winsock1.SendData strData
'


'End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
'Dim CardReturn As String
'CardReturn = ">513 0 2 3 0 17 2 A12 502 AC0 3 B241 Y2 ID1 Y3 PPF118 B50" & Chr(10) & "D1" & Chr(10) & "G27406871" & Chr(10) & "Q72499919" & Chr(10) & "b0A228E4926A30EFB" & Chr(10) & "d123456     3" & Chr(10) & "h0010010010190" & Chr(10) & "q;4519028522603304=23091204440300000201?" & Chr(10) & "U0" & Chr(10) & "3 PPH23 Z.37TELUSD01        00 2 T239 ;4519028522603304=23091204440300000201?2 TC2 012 TI8 TCMTCC165 api258 000000005 app258 000000004 ds2516 00000000000000004 dt2512 0410130954384 mt252 014 rf258 000001158 src_ex2516 00000000000000004 st252 000 9 3 ECP1 y2 HN13 206.87.131.272 HP4 65309 ISALLEGRO1 Y3 ISD1 Y2 MC1 Y3 PPP1 13 TID8 TCMTCC162 TO2 30"

Winsock1.GetData strData
'If strCommand = "EVNT" Then
'        MsgBox InStr(1, strData, "EVENTCODE:")
'        MsgBox InStr(1, strData, "LOCATION:")
'
'        Text1(8).Text = Mid(strData, InStr(1, strData, "EVENTCODE:") + 10, (InStr(1, strData, "LOCATION") - 1) - (InStr(1, strData, "EVENTCODE:") + 10))
'        If Len(Text1(8).Text) < 19 Then
'            Text1(8).Text = Trim(Text1(8).Text) & Space(19 - Len(Trim(Text1(8).Text)))
'        Else
'            Text1(8).Text = Left(Text1(8).Text, 19)
'        End If
'ElseIf strCommand = "LOGI" Then
'        Winsock1.GetData strData
'        Text1(8).Text = strData
'        MsgBox "LOGIN:" & Chr(10) & strData
'        strCommand = "DBT"
'        MsgBox "DBIT P" & CardReturn
'        Winsock1.SendData "DBIT P" & CardReturn
'ElseIf strCommand = "DBT" Then
'    Winsock1.GetData strData
'    MsgBox "DBT: " & strData
'    Dsintx1.ParseTransaction strData
'    MsgBox "ResponseCode:" & Dsintx1.GetTfield("Response_Code")
'    Dsintx1.ConfirmTransaction
'Else
    Text1(8).Text = strData
'End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description
End Sub


