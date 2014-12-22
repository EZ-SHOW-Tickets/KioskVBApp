Attribute VB_Name = "MiTrade"
'''DEVICE
Global IP As String
Global TerminalID As String
Global strComputerName As String
Global intComPort As Integer
Global SessionCounter As String

'''DATABASE
Global wrkjet As Workspace
Global db As Database
Global rsShows As Recordset
Global rsShow As Recordset
Global rsTicketTypes As Recordset

'''TICKET INFORMATION
Global TicketPrinter As Integer   '1 BOCA   2 PA
Global SelectedTicket As Integer
Global TicketPrices() As Currency
Global TicketTypes()
Global TotalTicketSold() As Integer
Global BarCodeType As String
Global TotTicketCount As Integer
Global TicketCount As String
Global TicketStock As Integer

'''SHOW INFORMATION
Global ShowID As Integer
Global SQLShowID As Integer
Global ShowImage As String

'''CREDIT CARD INFORMATION
Global AccountNumber As String      'Track2 Data
Global Track1 As String
Global CardType As String
Global CardNumber As String
Global CardExpire As String
Global CardName As String
Global MOP As String                'Credit Card Type
Global TotalPrice As Currency
Global ZipCode As String
Global strDebitString As String     'For Debit Card Processing
Global strPageImageCC As String     'Suffix for page graphic indicating CC used
Global ConvenienceFee As Currency
Global ConvenienceFeeType As Integer
Global CouponCode As String
Global EMail As String

'''PROCESSING INFORMATION
Global SwipTime As Date
Global SessionID As Long
Global webCounter As Integer
Global iSwipe As Integer
Global PrinterError As Boolean
Global SwipeProcess As Boolean

'''TICKETMASTER
Global SellerCode As String
Global Password As String
Global JournalID As String
Global JComputerID As String
Global EventCodeID As String
Global EventType As String
Global EventCity As String
Global EventReference As String
Global EventID As String
Global EventCode As String
Global ShowReference As String
Global EPDATE As String
Global bServer As Boolean           'TICKETMASTER Server does not respond
Global bOffLine As Boolean          'Off-Line Mode

'''AUTHORIZENET
Global bAN As Boolean               'Use AuthorizeNet Indicator
Global ANResponseCode As Integer
Global ANTransactionCode As String

'''CONTROLS
Global bChange As Boolean           'Change Purchase Indicator
Global dbLocal As Database          '***NOT USED
Global bTest As Boolean             'Use Test Event - **NOT USED**
Global bDebug As Boolean            'Debug messages provided
Global bDebit As Boolean            'Indicates a Debit Card is being used
Global bMultiple As Boolean         'Multiple event indicator
Global bNoPorts As Boolean          '***NOT USED
Global bDEMO As Boolean             'Credit Card Authorization bypassed
Global bTicketmation As Boolean     'Indicated a single DEMO is executing
Global bStartup As Boolean          'Indicates on STARTUP Screen
Global NotAllowDebit As Boolean     'NO debit cards (always TRUE)
Global VisaMCOnly As Boolean        'AMEX Allowed (True-NO, False-YES)
Global AllowDiscover As Boolean     'Allow Discover
Global GetZip As Boolean            'Acquire Zip Code
Global bGetEMail As Boolean          'Acquire EMail
Global bRestart As Boolean
Global NoAN As Boolean              'Do not actually Proceess Card
Global DBOpen As Boolean            'Indicates if DB has been open
Global bConvenienceFee As Boolean   'Charge a Convenience Fee
Global bAcceptConvenienceFee As Boolean
Global bReceipt As Boolean          'Ask about Receipt
Global bPrintReceipt As Boolean     'Print Receipt
Global bNoPrintBypass As Boolean    'Don't print Ticlets
Global UseAOP As Boolean            'AOP Screen
Global AOPBarcode As Boolean        'Use AOP Barcodes
Global bBatch As Boolean            'Uploading Batch info
Global nBatch As Integer            'Number of Batch items
Global bUseBU As Boolean            'USE backup website
Global iMaxTix As Integer           'Maximum tickets until ZIP Requested
Global bPrintError As Boolean       'Indicates that restart is from print error

Private Declare Function GetComputerName _
Lib "kernel32" Alias "GetComputerNameA" ( _
ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH   As Long = 15&

Public Declare Function GetInputState Lib "user32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


 'API's Function Declarations

 Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long


 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long


 Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long


 Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As Any, _
    ByVal lpWindowName As String) As Long

Public Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

 'API Constants

 Public Const GWL_STYLE = -16

 Public Const WS_DISABLED = &H8000000

 Public Const WM_CANCELMODE = &H1F

 Public Const WM_CLOSE = &H10


 Public Function IsTaskRunning(sWindowName As String) As Boolean
    Dim hwnd As Long, hWndOffline As Long
   
    On Error GoTo IsTaskRunning_Eh
    'get handle of the application
    'if handle is 0 the application is currently not running
    hwnd = FindWindow(0&, sWindowName)
    If hwnd = 0 Then
        IsTaskRunning = False
        Exit Function
    Else
        IsTaskRunning = True
    End If
   

IsTaskRunning_Exit:
        Exit Function


IsTaskRunning_Eh:
    Call ShowError(sWindowName, "IsTaskRunning")

 End Function


 Public Function EndTask(sWindowName As String) As Integer
    Dim x As Long, ReturnVal As Long, TargetHwnd As Long
   
    'find handle of the application
    TargetHwnd = FindWindow(0&, sWindowName)
    If TargetHwnd = 0 Then Exit Function
   
    If IsWindow(TargetHwnd) = False Then
        GoTo EndTaskFail
    Else
    'close application
        If Not (GetWindowLong(TargetHwnd, GWL_STYLE) And WS_DISABLED) Then
            x = PostMessage(TargetHwnd, WM_CLOSE, 0, 0&)
            DoEvents
        End If
    End If
   
    GoTo EndTaskSucceed


EndTaskFail:
    ReturnVal = False
    MsgBox "EndTask: cannot terminate " & sWindowName & " task"
    GoTo EndTaskEndSub


EndTaskSucceed:
    ReturnVal = True


EndTaskEndSub:
    EndTask% = ReturnVal

 End Function


 Public Function ShowError(sText As String, sProcName As String)
   'this function displays an error that occurred
   
    Dim sMsg As String
    sMsg = "Error # " & str(Err.Number) & " was generated by " _
         & Err.Source & vbCrLf & Err.Description
    MsgBox sMsg, vbCritical, sText & Space(1) & sProcName
    Exit Function


 End Function

Public Function CurrentMachineName() As String

Dim lSize As Long
Dim sBuffer As String
sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
lSize = Len(sBuffer)

If GetComputerName(sBuffer, lSize) Then
    CurrentMachineName = Left$(sBuffer, lSize)
End If

End Function
Sub LogClick(FromForm, Click)
db.Execute "Insert into LOG (LogTime,SessionID,ClickFrom) values(#" & Now() & "#," & SessionID & ",'" & FromForm & ":" & Click & "')"

End Sub

Sub OpenDataBase()
    Dim rsInfo As Recordset
'    On Error GoTo DB_Error
    Set wrkjet = CreateWorkspace("", "admin", "", dbUseJet)
    Dim strpath As String
    strpath = App.path & "\WUS.mdb"
    Set db = wrkjet.OpenDataBase(strpath)
'    Set dbLocal = wrkjet.OpenDataBase(strpath)
'    Set rsInfo = db.OpenRecordset("Select * from SETUP_INFO")
'    If rsInfo.EOF Then
'        dbLocal.Execute "INSERT INTO SETUP_INFO (ComputerName,DatabaseServer) values('" & CurrentMachineName() & "','" & App.Path & "')"
'        Set db = dbLocal
'    Else
'        'If StrComp(App.Path, rsInfo("DatabaseServer")) <> 0 Then
'           'different server
'            strComputerName = rsInfo("ComputerName")
'            strpath = rsInfo("DatabaseServer") & "\WUS.mdb"
'            Set db = wrkjet.OpenDataBase(strpath)
'        'End If
'    End If
'DB_Error:
'    dbLocal.Execute "DELETE from SETUP_INFO"
'    dbLocal.Execute "INSERT INTO SETUP_INFO (ComputerName,DatabaseServer) values('" & CurrentMachineName() & "','" & App.Path & "')"
'    Set rsInfo = dbLocal.OpenRecordset("Select * from SETUP_INFO")
'    strpath = rsInfo("DatabaseServer") & "\WUS.mdb"
'    Set db = wrkjet.OpenDataBase(strpath)
End Sub


