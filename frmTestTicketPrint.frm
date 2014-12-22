VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTestTicketPrint 
   Caption         =   "TicketPrint"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerACK 
      Left            =   120
      Top             =   1440
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Print Text"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton DMPrint 
      Caption         =   "DataMax Receipt"
      Height          =   975
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton DataMaxPrintTicket 
      Caption         =   "Practical Automation"
      Height          =   975
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdCoupon 
      Caption         =   "PRINT Coupon"
      Height          =   915
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   1515
   End
   Begin VB.TextBox txtNoTix 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1125
      TabIndex        =   1
      Text            =   "1"
      Top             =   225
      Width           =   1005
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   195
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT TICKET"
      Height          =   915
      Left            =   720
      TabIndex        =   0
      Top             =   795
      Width           =   1755
   End
End
Attribute VB_Name = "frmTestTicketPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub PrintTicketsBocaT(i As Integer, K As Integer)
Dim strPrint As String
Dim printLocation As String
Dim iPrintLocation As Integer
Dim printShow As String
Dim iPrintShow As Integer
Dim printPrice As String
'Dim testJournalID As Long
'Dim testDeviceID As String
Dim iTicketType As String
Dim rsHeaders As Recordset
Dim rsBarCodes As Recordset
Dim thisType As Integer
Dim iCycles As Integer
Dim strBarCode As String
Dim strBarCodeTXT As String
Dim strEventType As String
Dim printTax1 As String

'printTax1 = "$12.86+1.14tax"
On Error GoTo PrintErr

If bDebug Then MsgBox "ENTER HSG PRINT"
iCycles = 0
'Get Show information
Set rsHeaders = db.OpenRecordset("Select * from TICKET_HEADERS where ShowID=" & ShowID)
'Set EVENT TYPE for ticket
    EventType = "AUTO SHOW"
'Create appropriate Barcode
Select Case TicketTypes(1, i)
    Case 1
        iTicketType = "81"
        'printTax1 = "$12.86+$1.14tax"
    Case 2
        iTicketType = "84"
    Case 3
        iTicketType = "82"
        'printTax1 = "$3.67+$0.33tax"
    Case 16
        iTicketType = "83"
        'printTax1 = "$4.59+$0.41tax"
    Case Else
        iTicketType = "85"
End Select
strBarCode = iTicketType & Format(IP, "0#") & Trim(str(SessionID)) & Format(K, "0#")
strBarCodeTXT = strBarCode

'Show Location
printLocation = rsHeaders("Location")
printShow = rsHeaders("Header2")
printPrice = Format(TicketTypes(3, i), "$#.00")
'Calculate print locations for SHOW and VENUE
iPrintShow = 395 - CInt((Len(printShow) * 25) / 2)
iPrintLocation = 410 - CInt((Len(printLocation) * 20) / 2)

    'BOCA PRINTER
    strPrint = "<NR><RC50," & CStr(iPrintLocation) & "><F8><HW1,1>" & printLocation
   'Move price up to same level as Location
    strPrint = strPrint & "<NR><RC80,15><F7><HW1,1>" & printPrice
    'Add tax
    'strPrint = strPrint & "<NR><RC105,15><F2><HW1,1>" & printTax1
    ' was strPrint = strPrint & "<NR><RC90,15><F7><HW1,1>" & printPrice
    strPrint = strPrint & "<RC80,700><F7>" & printPrice
    'Add tax
    'strPrint = strPrint & "<NR><RC105,700><F2><HW1,1>" & printTax1
    strPrint = strPrint & "<RC130,15><F7>" & rsHeaders("Header5")
    strPrint = strPrint & "<RC120,150><LT3><HX525>"
    strPrint = strPrint & "<RC130," & CStr(iPrintShow) & "><F10>" & printShow
    strPrint = strPrint & "<RC130,700><F7>" & rsHeaders("Header5")
    strPrint = strPrint & "<RC180,15><F3><HW2,1>" & TicketTypes(5, i)
    strPrint = strPrint & "<RC180,150><F10><HW1,1>**** " & EventType & " ****"
    strPrint = strPrint & "<RC180,700><F3><HW2,1>" & TicketTypes(5, i)
    strPrint = strPrint & "<RC220,150><LT3><HX525>"
    strPrint = strPrint & "<RC270,15><F3><HW1,1>" & CStr(K) & " of " & TotTicketCount
    strPrint = strPrint & "<RC230,200><F3>" & rsHeaders("Header4")
    strPrint = strPrint & "<RC270,700><F3>" & CStr(K) & " of " & TotTicketCount
    strPrint = strPrint & "<RC270,250><F9>" & Now()
    strPrint = strPrint & "<RC290,200><F13><HW2,1>NO REFUNDS OR EXCHANGES"
    strPrint = strPrint & "<RC320,15><F9><HW1,1>" & Format(Now(), "MM/DD/YYYY")
    strPrint = strPrint & "<RC320,700><F9><HW1,1>" & Format(Now(), "MM/DD/YYYY")
    strPrint = strPrint & "<RC0,850><LT5><HY370>"
    strPrint = strPrint & "<RC40,1005><LT5><VX330>"
    strPrint = strPrint & "<RC40,1035><LT5><VX330>"
    strPrint = strPrint & "<RL><RC300,870><F3><HW1,1>" & strBarCodeTXT

    strPrint = strPrint & "<RC60,990><FL10><X3>:" & strBarCode & ":"
  
    If bDebug Then MsgBox "DISCLAIMER"

    'Disclaimer
    iPrintShow = 200 + CInt((Len(printShow & " " & EventType) * 13) / 2)
    strPrint = strPrint & "<RC" & iPrintShow & ",1015><F9><HW1,1>" & printShow & " " & EventType
    strPrint = strPrint & "<RC360,1045><F2><HW1,1>Taxes Incl. Subject to Show Rules"
    strPrint = strPrint & "<RC270,1060><F2><HW1,1>" & ANTransactionCode
    If bDebug Then MsgBox "Send Print Command"
    strPrint = strPrint & "<p>"

    MSComm1.output = strPrint
    
    ack_found = False
    If wait_status = True Then
        wait_for_ack                 'Wait for the ACK
    End If
    If problem = True Then GoTo PrintErr

'''''REPLACE WITH ACK
'    Sleep 700
'   Do
'        DoEvents
'        Buffer$ = Buffer$ & MSComm1.Input
'        iCycles = iCycles + 1
'        If iCycles > 10 Then
'        'If iCycles > 30 Then
'           PrinterError = True
'           If bDebug Then MsgBox "PrinterError Set, iCycles=" & CStr(iCycles)
'
'            Buffer$ = "XX"
'        End If
'    Loop Until Len(Buffer$) > 0
'If bDebug Then MsgBox "End Buffer Loop"
''''''''''''''''''''''
Call LogClick("Print BarCode", strBarCode)
Exit Sub
PrintErr:
MsgBox Err.Description

End Sub



Private Sub cmdCoupon_Click()
Dim strPrint As String
Dim printLocation As String
Dim iPrintLocation As Integer
Dim printShow As String
Dim iPrintShow As Integer

Dim testJournalID As Long
Dim testDeviceID As String
Dim testTicketType As String
Dim testNumTix As Integer
Dim i As Integer

MSComm1.CommPort = 4
MSComm1.Settings = "9600,N,8,1"
MSComm1.OutBufferSize = 1024
MSComm1.InputLen = 0
MSComm1.PortOpen = True
'<SP#,#><pcx><G#>pcx bytes
'Dim InputData
'Open App.Path & "/HSG.pcx" For Input As #1 'Open file for input.
'Do While Not EOF(1) 'Check for end of file.
'    Line Input #1, InputData 'Read line of data.
'InputData = Input(LOF(1), 1)
'     Input #1, InputData
'    Debug.Print InputData 'Print to the Immediate window.
'Loop
'Close #1 'Close file.

Open App.path & "/car.pcx" For Binary As #1
Dim strBuff As String
strBuff = Space(LOF(1))
Get #1, , strBuff
Close #1

    strPrint = "<SP50,100><pcx><G6704>" & strBuff
    'strPrint = "<SP50,100><pcx><G2497>" & strBuff
    strPrint = strPrint & "<p>"
    MSComm1.output = strPrint
'
'       Sleep 700
'
'   Do
'      DoEvents
'   buffer$ = buffer$ & MSComm1.Input
'   Loop Until Len(buffer$) > 0
'
   MSComm1.PortOpen = False
End Sub

Private Sub cmdPrint_Click()
Dim strPrint As String
Dim printLocation As String
Dim iPrintLocation As Integer
Dim printShow As String
Dim iPrintShow As Integer
Dim TestBarCode As String
Dim TestCodes(10) As String

Dim testJournalID As Long
Dim testDeviceID As String
Dim testTicketType As String
Dim testNumTix As Integer
Dim i As Integer

'''''''''''''
ShowID = 70
ReDim TicketTypes(6, 1)

TicketTypes(1, 1) = 1
TicketTypes(3, 1) = 10#
TicketTypes(5, 1) = "ADULT"
TotTicketCount = 1
ANTransactionCode = "333333333"

wait_status = True
Response = True
ready = True
problem = False
duration = 0.5
Display_Flag = True



MSComm1.CommPort = 4
MSComm1.Settings = "9600,N,8,1"
MSComm1.OutBufferSize = 1024
MSComm1.InputLen = 0
MSComm1.PortOpen = True
MSComm1.RThreshold = 1
MSComm1.SThreshold = 1
MSComm1.InputMode = comInputModeText
MSComm1.InputLen = 0
MSComm1.InBufferCount = 0
MSComm1.InputMode = comInputModeText
MSComm1.Handshaking = comNone
'Call PrintTicketsBocaT(1, 1)
'Call PrintTicketsHSG(1, 1)
   
printLocation = "LOS ANGELES CONV CENTER"
'printLocation = "JACOB K JAVITS CONV CENTER"
printShow = "2009 LA INTL"
'printShow = "2009 NY INTL"
testJournalID = 123456
testDeviceID = "19"
testTicketType = "811"
testNumTix = CInt(txtNoTix.text)
TestCodes(1) = "SWlw72TSywzZ"
TestCodes(2) = "TTGhkQV8ZY9c"
TestCodes(3) = "D6XkJDF2wYMc"
TestCodes(4) = "lXRxVDm8WMMp"
TestCodes(5) = "K5lxVDL8YlmZ"
TestCodes(6) = "d17wJQf5M9mc"
TestCodes(7) = "R6WwJQS8398N"


TestBarCode = "SWlw72TSywzZ"
iPrintShow = 410 - CInt((Len(printShow) * 25) / 2)
iPrintLocation = 410 - CInt((Len(printLocation) * 20) / 2)

For i = 1 To testNumTix
    TestBarCode = TestCodes(i)
    strPrint = "<NR><RC50," & CStr(iPrintLocation) & "><F8>" & printLocation
    strPrint = strPrint & "<NR><RC90,10><F7>$10.00"
    strPrint = strPrint & "<RC90,700><F7>$10.00"
    'strPrint = strPrint & "<RC130,10><F7>2008NY"
    strPrint = strPrint & "<RC130,10><F7>2008LA"
    strPrint = strPrint & "<RC120,150><LT3><HX525>"
    strPrint = strPrint & "<RC130," & CStr(iPrintShow) & "><F10>" & printShow
    'strPrint = strPrint & "<RC130,700><F7>2008NY"
    strPrint = strPrint & "<RC130,700><F7>2008LA"
    strPrint = strPrint & "<RC180,10><F3><HW2,1>ADULT"
    strPrint = strPrint & "<RC180,150><F10><HW1,1>**** AUTO SHOW****"
    strPrint = strPrint & "<RC180,700><F3><HW2,1>ADULT"
    strPrint = strPrint & "<RC220,150><LT3><HX525>"
    strPrint = strPrint & "<RC270,10><F3><HW1,1>" & CStr(i) & " of " & testNumTix
    strPrint = strPrint & "<RC230,200><F3>Dec 4 to 13, 2009"
    strPrint = strPrint & "<RC270,700><F3>" & CStr(i) & " of " & testNumTix
    strPrint = strPrint & "<RC270,250><F9>" & Now()

    strPrint = strPrint & "<RC300,200><F13><HW2,1>GOOD FOR ONE DAY ONLY"
    strPrint = strPrint & "<RC40,1005><LT5><VX330>"
    strPrint = strPrint & "<RC40,1035><LT5><VX330>"
    'ABOVE BAR CODE
    'strPrint = strPrint & "<RL><RC320,870><F3><HW1,1>" & testTicketType & testDeviceID & CStr(testJournalID) & CStr(i)
    strPrint = strPrint & "<RL><RC320,870><F3><HW1,1>" & TestBarCode
    'BAR CODE
    'strPrint = strPrint & "<RC20,990><NL10><X2>*" & testTicketType & testDeviceID & CStr(testJournalID) & CStr(i) & "*"
    strPrint = strPrint & "<RC20,990><OL10><X2>^" & TestBarCode & "^"
    'Disclaimer
    'strPrint = strPrint & "<RC360,1015><F9><HW1,1>2009 NY INTL AUTO SHOW"
    strPrint = strPrint & "<RC360,1015><F9><HW1,1>2009 LA INTL AUTO SHOW"
    strPrint = strPrint & "<RC360,1045><F2><HW1,1>Taxes Incl. Subject to Show Rules"
    strPrint = strPrint & "<RC340,1060><F2><HW1,1>No Refunds  No Exchanges"
    'strPrint = strPrint & "<RC380,1075><F2><HW1,1>Holder releases SHOW for any liability"

    strPrint = strPrint & "<p>"
    MSComm1.output = strPrint

       Sleep 700

'   Do
'      DoEvents
'   Buffer$ = Buffer$ & MSComm1.Input
'   Loop Until Len(Buffer$) > 0
   Next i
'
 MSComm1.PortOpen = False

End Sub


Private Sub Command1_Click()


End Sub


Private Sub cmdText_Click()
Dim strPrint As String
Dim printLocation As String
Dim iPrintLocation As Integer
Dim printShow As String
Dim iPrintShow As Integer
Dim TestBarCode As String
Dim TestCodes(10) As String

Dim testJournalID As Long
Dim testDeviceID As String
Dim testTicketType As String
Dim testNumTix As Integer
Dim i As Integer
Dim CR, LF As String
    CR = Chr$(13)       'Use this for a carriage return
    LF = Chr$(10)       'Use this for a line feed


MSComm1.CommPort = 4
MSComm1.Settings = "9600,N,8,1"
MSComm1.OutBufferSize = 1024
MSComm1.InputLen = 0
MSComm1.PortOpen = True
   
printLocation = "               LOS ANGELES CONV CENTER" ' & CR & LF
'printLocation = "JACOB K JAVITS CONV CENTER"
printShow = "                    2009 LA INTL"
'printShow = "2009 NY INTL"
testJournalID = 123456
testDeviceID = "19"
testTicketType = "811"
testNumTix = CInt(txtNoTix.text)
TestCodes(1) = "SWlw72TSywzZ"
TestCodes(2) = "TTGhkQV8ZY9c"
TestCodes(3) = "D6XkJDF2wYMc"
TestCodes(4) = "lXRxVDm8WMMp"
TestCodes(5) = "K5lxVDL8YlmZ"
TestCodes(6) = "d17wJQf5M9mc"
TestCodes(7) = "R6WwJQS8398N"


TestBarCode = "SWlw72TSywzZ"
iPrintShow = 410 - CInt((Len(printShow) * 25) / 2)
iPrintLocation = 410 - CInt((Len(printLocation) * 20) / 2)

For i = 1 To testNumTix
    TestBarCode = TestCodes(i)
    strPrint = printLocation
    strPrint = printShow

    strPrint = strPrint & "<p>"
    MSComm1.output = strPrint
  
       Sleep 700

'   Do
'      DoEvents
'   Buffer$ = Buffer$ & MSComm1.Input
'   Loop Until Len(Buffer$) > 0
   Next i
   
   MSComm1.PortOpen = False

End Sub

Private Sub DataMaxPrintTicket_Click(Index As Integer)
Dim strPrint As String
Dim printLocation As String
Dim iPrintLocation As Integer
Dim printShow As String
Dim iPrintShow As Integer
Dim TestBarCode As String
Dim TestCodes(10) As String

Dim testJournalID As Long
Dim testDeviceID As String
Dim testTicketType As String
Dim testNumTix As Integer
Dim i As Integer
Dim p As VB.Printer

For Each p In VB.Printers
   If p.DeviceName = "Generic / Text Only" Then
        Set Printer = p
   End If
Next


printLocation = "LOS ANGELES CONV CENTER"
'printLocation = "JACOB K JAVITS CONV CENTER"
printShow = "2009 LA INTL"
'printShow = "2009 NY INTL"
testJournalID = 123456
testDeviceID = "19"
testTicketType = "811"
testNumTix = CInt(txtNoTix.text)
TestCodes(1) = "SWlw72TSywzZ"
TestCodes(2) = "TTGhkQV8ZY9c"
TestCodes(3) = "D6XkJDF2wYMc"
TestCodes(4) = "lXRxVDm8WMMp"
TestCodes(5) = "K5lxVDL8YlmZ"
TestCodes(6) = "d17wJQf5M9mc"
TestCodes(7) = "R6WwJQS8398N"


TestBarCode = "SWlw72TSywzZ"
iPrintShow = 410 - CInt((Len(printShow) * 25) / 2)
iPrintLocation = 410 - CInt((Len(printLocation) * 20) / 2)

i = 1
    TestBarCode = TestCodes(i)
    Printer.Print "<NR><RC50," & CStr(iPrintLocation) & "><F8>" & printLocation
    Printer.Print "<NR><RC90,10><F7>$10.00"
    Printer.Print "<RC90,700><F7>$10.00"
    Printer.Print "<RC130,10><F7>2008LA"
'''    strPrint = "<NR><RC50," & CStr(iPrintLocation) & "><F8>" & printLocation
'''    strPrint = strPrint & "<NR><RC90,10><F7>$10.00"
'''    strPrint = strPrint & "<RC90,700><F7>$10.00"
'''    strPrint = strPrint & "<RC130,10><F7>2008LA"
    Printer.Print "<RC120,150><LT3><HX525>"
    Printer.Print "<RC130," & CStr(iPrintShow) & "><F10>" & printShow
    Printer.Print "<RC130,700><F7>2008LA"
    Printer.Print "<RC180,10><F3><HW2,1>ADULT"
    Printer.Print "<RC180,150><F10><HW1,1>**** AUTO SHOW****"
    Printer.Print "<RC180,700><F3><HW2,1>ADULT"
    Printer.Print "<RC220,150><LT3><HX525>"
    Printer.Print "<RC270,10><F3><HW1,1>" & CStr(i) & " of " & testNumTix
    Printer.Print "<RC230,200><F3>Dec 4 to 13, 2009"
    Printer.Print "<RC270,700><F3>" & CStr(i) & " of " & testNumTix
    Printer.Print "<RC270,250><F9>" & Now()

    Printer.Print "<RC300,200><F13><HW2,1>GOOD FOR ONE DAY ONLY"
    Printer.Print "<RC10,1005><LT5><VX330>"
    Printer.Print "<RC10,1035><LT5><VX330>"
    'ABOVE BAR CODE
    'printer.print "<RL><RC320,870><F3><HW1,1>" & testTicketType & testDeviceID & CStr(testJournalID) & CStr(i)
    Printer.Print "<RL><RC320,870><F3><HW1,1>" & TestBarCode
    'BAR CODE
    'printer.print "<RC20,990><NL10><X2>*" & testTicketType & testDeviceID & CStr(testJournalID) & CStr(i) & "*"
    Printer.Print "<RC20,990><OL10><X2>^" & TestBarCode & "^"
    'Disclaimer
    'printer.print "<RC360,1015><F9><HW1,1>2009 NY INTL AUTO SHOW"
    Printer.Print "<RC340,1015><F9><HW1,1>2009 LA INTL AUTO SHOW"
    Printer.Print "<RC340,1045><F2><HW1,1>Taxes Incl. Subject to Show Rules"
    Printer.Print "<RC320,1060><F2><HW1,1>No Refunds  No Exchanges"
    'printer.print "<RC380,1075><F2><HW1,1>Holder releases SHOW for any liability"

    Printer.Print "<p>"
    'Printer.Print strPrint
    Printer.EndDoc
End Sub

Private Sub DMPrint_Click(Index As Integer)
    Dim p As VB.Printer
    EventType = "FAIR"
    printLocation = "Location"
    printShow = "2010 ALAMEDA COUNTY"
    printPrice = Format("20.00", "$#.00")
    TotTicketCount = 10
    iPrintShow = 410 - CInt((Len(printShow) * 25) / 2)
    iPrintLocation = 410 - CInt((Len(printLocation) * 20) / 2)
    ANTransactionCode = "1234567890"
    MOP = "MC"
    CardNumber = "12345"
    strBarCodeTXT = "SWlw72TSywzZ"
    For Each p In VB.Printers
        'If p.DeviceName = "PA ITL2002F PT" Then
        If p.DeviceName = "Generic / Text Only" Then
        Set Printer = p
        End If
    Next

    Printer.Print "<NR><RC20," & CStr(iPrintLocation) & "><F8><HW1,1>" & printLocation
    Printer.Print "<NR><RC60,15><F7><HW1,1>" & printPrice
    Printer.Print "<RC60,700><F7>" & printPrice
   ' Printer.Print "<RC100,15><F7>" & "2010ACF"
    Printer.Print "<RC90,150><LT3><HX525>"
    Printer.Print "<RC100," & CStr(iPrintShow) & "><F10>" & printShow
    'Printer.Print "<RC100,700><F7>" & rsHeaders; "2010ACF"
    Printer.Print "<RC150,150><F10><HW1,1>**** " & EventType & " ****"
    Printer.Print "<RC190,150><LT3><HX525>"
    If TotTicketCount > 1 Then
        Printer.Print "<RC240,15><F3>" & TotTicketCount & " Tickets"
    Else
        Printer.Print "<RC240,15><F3>" & TotTicketCount & " Ticket"
    End If
    Printer.Print "<RC200,200><F3>" & MOP & "  ****" & Right(CardNumber, 4)
    If TotTicketCount > 1 Then
        Printer.Print "<RC240,700><F3>" & TotTicketCount & " Tickets"
    Else
        Printer.Print "<RC240,700><F3>" & TotTicketCount & " Ticket"
    End If
    Printer.Print "<RC240,250><F9>" & Now()
    Printer.Print "<RC270,200><F13><HW2,1>***** RECEIPT ****"

    
    Printer.Print "<RC290,15><F9><HW1,1>" & Format(Now(), "MM/DD/YYYY")
    Printer.Print "<RC290,700><F9><HW1,1>" & Format(Now(), "MM/DD/YYYY")
    Printer.Print "<RC0,850><LT5><HY370>"
    Printer.Print "<RC10,1005><LT5><VX330>"
    Printer.Print "<RL><RC270,910><F9><HW1,1>" & strBarCodeTXT
    Printer.Print "<RC340,970><F10><HW1,1>**RECEIPT**"

    Printer.Print "<RC340,1045><F2><HW1,1>KEEP THIS RECEIPT FOR YOUR RECORDS"
    Printer.Print "<RC240,1080><F2><HW1,1>" & ANTransactionCode
    Printer.Print "<p>"

    Printer.EndDoc

End Sub


Private Sub MSComm1_OnComm()
    'Possible Errors or events control through case statement
    'Handle each event or error by placing code below each case statement
    'Comment out the Display_Text statements not needed.
    Select Case MSComm1.CommEvent

        'A Break was received
        Case comEventBreak
            Call Display_Text("A Break was received.")             'debug statement
    
        'CD (RLSD) Timeout
        Case comEventCDTO
            Call Display_Text("CD (RLSD) Timeout.")                'debug statement
    
        'CTS Timeout
        Case comEventCTSTO
            Call Display_Text("CTS Timeout.")                      'debug statement
    
        'DSR Timeout
        Case comEventDSRTO
            Call Display_Text("DSR Timeout.")                      'debug statement
    
        'Framing Error
        Case comEventFrame
            Call Display_Text("Framing Error.")                    'debug statement
    
        'Data Lost
        Case comEventOverrun
            Call Display_Text("Data Lost.")                        'debug statement
    
        'Receive buffer overflow
        Case comEventRxOver
           Call Display_Text("Receive buffer overflow.")          'debug statement
    
        'Parity Error
        Case comEventRxParity
           Call Display_Text("Parity Error.")                     'debug statement
            
        'Transmit buffer full
        Case comEventTxFull
            Call Display_Text("Transmit buffer full.")             'debug statement
            
        'Unexpected error retrieving DCB]
        Case comEventDCB
            Call Display_Text("Unexpected error retrieving DCB.")  'debug statement

        'Change in the CD line.
        Case comEvCD
           Call Display_Text("Change in the CD line. CD = " & ComPort.CDHolding)      'debug statement
'            If (ComPort.CDHolding = True) Then
'                ShpCD.FillColor = Green
'            Else
'                ShpCD.FillColor = Red
'            End If
        
        'Change in the CTS line.
        Case comEvCTS
           Call Display_Text("Change in the CTS line. CTS = " & ComPort.CTSHolding)      'debug statement
'            If (ComPort.CTSHolding = True) Then
'                ShpCTS.FillColor = Green
'            Else
'                ShpCTS.FillColor = Red
'            End If
        
        'Change in the DSR line.
        Case comEvDSR
            Call Display_Text("Change in the DSR line. DSR = " & ComPort.DSRHolding)      'debug statement
'            If (ComPort.DSRHolding = True) Then
'                ShpDSR.FillColor = Green
'            Else
'                ShpDSR.FillColor = Red
'            End If
        
        'Change in the Ring Indicator.
        Case comEvRing
            Call Display_Text("Change in the Ring Indicator.")      'debug statement
    
        'Received RThreshold number of chars.
        Case comEvReceive
            Call Display_Text("Receive. Ready = " & ready)         'debug statement
'            ShpRD.FillColor = Green                 'make Receive data light green
'            ShpTD.FillColor = Red                   'make Transmit data light red
        
            'The ready flag is set in find_baud() once the baud rate has
            'been established.  From that point forward all reading of data
            'from the printer will happen through here
            'If (ready) Then
            '    'Buffer = Buffer & FrmBoca.ComPort.Input
            '    Receive (MAX_BUFFER_SIZE)           'read print buffer
            'End If
            If (ready) Then
                Buffer = Buffer & MSComm1.Input
                Receive (MAX_BUFFER_SIZE)           'read print buffer
            End If

        'There are SThreshold number of characters in the transmit buffer.
        Case comEvSend
            Call Display_Text("Transmit")          'debug statement
'            ShpRD.FillColor = Red                   'make Receive data light red
'            ShpTD.FillColor = Green                 'make Transmit data light green
    
        'An EOF charater was found in the input stream
        Case comEvEOF
            Call Display_Text("EOF")               'debug statement
    
        'Report possible error condition for unknown comm event
        Case Else
            Call Display_Text("Unknown ComPort case = " & MSComm1.CommEvent)
        
    End Select

End Sub

Private Sub TimerACK_Timer()
    timer_flag = True
    TimerACK.Enabled = False

End Sub


