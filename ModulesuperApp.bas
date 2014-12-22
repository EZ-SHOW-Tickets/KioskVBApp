Attribute VB_Name = "Module1"
Global Const strConnectionString = "DRIVER={SQL Server};SERVER=207.171.1.162;DATABASE=Ticketmation;UID=sa;PWD=cgc0023"
Global IP As String
Global wrkjet As Workspace
Global db As Database

'API Constants

 Public Const GWL_STYLE = -16

 Public Const WS_DISABLED = &H8000000

 Public Const WM_CANCELMODE = &H1F

 Public Const WM_CLOSE = &H10

 Public Const SMTO_BLOCK = &H1
 
 Public Const SMTO_ABORTIFHUNG = &H2
 
 Public Const WM_NULL = &H0

Public Const GW_HWNDNEXT = 2

Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public hWndApp As Long


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" _
    (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As _
    Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function OpenProcess Lib "Kernel32.dll" _
  (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
      ByVal dwProcId As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
 Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
 As String, ByVal lpReserved As Long, lpType As Long, lpData As _
 Any, lpcbData As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
 "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
 ByVal ulOptions As Long, ByVal samDesired As Long, _
 phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
 (ByVal hKey As Long) As Long

Public Const KEY_QUERY_VALUE = &H1

Public Const HKEY_DYN_DATA As Long = &H80000006
 Public Const STAT_DATA = "PerfStats\StatData"
 Public Const CPU_USE = "KERNEL\CPUUsage"


Private Declare Sub GlobalMemoryStatus Lib "kernel32" (ByRef lpBuffer As MemoryStatus)

Private Type MemoryStatus
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Sub DownLoadNewDB()
Dim errcode As Long
Dim url As String
Dim localFileName As String

url = "http://207.171.1.163/WUSUpdate.mdb"
localFileName = App.Path & "\WUS" & Format(Now, "mmddyy") & ".mdb"

errcode = URLDownloadToFile(0, url, localFileName, 0, 0)
If errcode = 0 Then
    MsgBox "Download ok"
    Name App.Path & "\WUS.mdb" As App.Path & "\WUSOld" & Format(Now, "mmddyy") & ".mdb"
    FileCopy App.Path & "\WUS" & Format(Now, "mmddyy") & ".mdb", App.Path & "\WUS.mdb"
    MsgBox "DB Update Complete"
Else
    MsgBox "Error while downloading"
End If

End Sub


Public Function GetMemoryLoad() As String
'::::::::::::::::::::::::::::::::::::::::::::::::::::::: METHOD / FUNCTION
':: DESCRIPTION:  Returns the memory load percentage.                   ::
'::       INPUT:  NONE                                                  ::
'::      OUTPUT:  NONE                                                  ::
':: ASSUMPTIONS:  NONE                                                  ::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: LOCAL VARIABLES:                                                    ::
'::---------------------------------------------------------------------::
    Dim memMemory As MemoryStatus
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    On Error GoTo ErrorHandler

    GlobalMemoryStatus memMemory                                ' GET STATUS
    GetMemoryLoad = memMemory.dwMemoryLoad & "% used"           ' RETURN PERCENT
    
    Exit Function
ErrorHandler:
    Err.Clear
    GetMemoryLoad = "N/A"
End Function

Public Function CPU_Usage_Percent() As Long
On Error GoTo ErrorHandler:
Dim lResult As Long
Dim lData As Long
Dim lType As Long
Dim hKey As Long


lResult = RegOpenKeyEx(HKEY_DYN_DATA, STAT_DATA, _
   0, KEY_QUERY_VALUE, hKey)

If lResult = 0 Then
    lResult = RegQueryValueEx(hKey, CPU_USE, 0, _
     lType, lData, 4)
    If lResult = 0 Then
        CPU_Usage_Percent = lData
        lResult = RegCloseKey(hKey)
    End If
End If
Exit Function

ErrorHandler:
     On Error Resume Next
     RegCloseKey hKey
     Exit Function
     End Function

'''''''
'
'
'
' Public Declare Function Process32First Lib "kernel32" ( _
'   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'
'Public Declare Function Process32Next Lib "kernel32" ( _
'   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
'
'Public Declare Function CloseHandle Lib "Kernel32.dll" _
'   (ByVal Handle As Long) As Long
'
'
'Public Declare Function EnumProcesses Lib "psapi.dll" _
'   (ByRef lpidProcess As Long, ByVal cb As Long, _
'      ByRef cbNeeded As Long) As Long
'
'Public Declare Function GetModuleFileNameExA Lib "psapi.dll" _
'   (ByVal hProcess As Long, ByVal hModule As Long, _
'      ByVal ModuleName As String, ByVal nSize As Long) As Long
'
'Public Declare Function EnumProcessModules Lib "psapi.dll" _
'   (ByVal hProcess As Long, ByRef lphModule As Long, _
'      ByVal cb As Long, ByRef cbNeeded As Long) As Long
'
'Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
'   ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'
'Public Declare Function GetVersionExA Lib "kernel32" _
'   (lpVersionInformation As OSVERSIONINFO) As Integer
'
'Public Type PROCESSENTRY32
'   dwSize As Long
'   cntUsage As Long
'   th32ProcessID As Long           ' This process
'   th32DefaultHeapID As Long
'   th32ModuleID As Long            ' Associated exe
'   cntThreads As Long
'   th32ParentProcessID As Long     ' This process's parent process
'   pcPriClassBase As Long          ' Base priority of process threads
'   dwFlags As Long
'   szExeFile As String * 260       ' MAX_PATH
'End Type
'
'Public Type OSVERSIONINFO
'   dwOSVersionInfoSize As Long
'   dwMajorVersion As Long
'   dwMinorVersion As Long
'   dwBuildNumber As Long
'   dwPlatformId As Long           '1 = Windows 95.
'                                  '2 = Windows NT
'
'   szCSDVersion As String * 128
'End Type
'
'Public Const PROCESS_QUERY_INFORMATION = 1024
'Public Const PROCESS_VM_READ = 16
'Public Const MAX_PATH = 260
'Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Public Const SYNCHRONIZE = &H100000
''STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
'Public Const PROCESS_ALL_ACCESS = &H1F0FFF
'Public Const TH32CS_SNAPPROCESS = &H2&
'Public Const hNull = 0
'
' Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'
'
' Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
'    ByVal hwnd As Long, _
'    ByVal nIndex As Long) As Long
'
'
' Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
'    ByVal hwnd As Long, _
'    ByVal wMsg As Long, _
'    ByVal wParam As Long, _
'    ByVal lParam As Long) As Long
'
'
'
''Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, _
''    lpdwProcessId As Long) As Long
'
''Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
''    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'
'
'
''Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
''    ByVal uExitCode As Long) As Long

Sub UpdateDB()
Dim i As Integer
Dim ShowID As String

Dim db2 As Database
Dim rsInfo As Recordset
Dim rsRecord As Recordset
Dim strpath As String
Dim strpath2 As String

Dim errcode As Long
Dim url As String
Dim localFileName As String

url = "http://207.171.1.163/WUSUpdate.mdb"
localFileName = App.Path & "\WUS" & Format(Now, "mmddyy") & ".mdb"

errcode = URLDownloadToFile(0, url, localFileName, 0, 0)
If errcode = 0 Then
    MsgBox "Download ok"
Set wrkjet = CreateWorkspace("", "admin", "", dbUseJet)
strpath = App.Path & "\WUS.mdb"
strpath2 = App.Path & "\WUS" & Format(Now, "mmddyy") & ".mdb"
Set db = wrkjet.OpenDatabase(strpath)
Set db2 = wrkjet.OpenDatabase(strpath2)
Set rsInfo = db2.OpenRecordset("Select * from UPDATE_WUS")

If rsInfo(0) < 0 Then
    'New Show... INSERT
        ShowID = CStr(-rsInfo(0))
        Set rsRecord = db2.OpenRecordset("Select * from SHOWS where ShowID=" & ShowID)
        db.Execute "INSERT INTO SHOWS values(" & rsRecord(0) & ",'" & rsRecord(1) & "',#" & rsRecord(2) & "#,#" & rsRecord(3) & "#,'" & rsRecord(4) & "','" & rsRecord(5) & "','" & rsRecord(6) & "'," & rsRecord(7) & "," & rsRecord(8) & "," & rsRecord(9) & ",'" & rsRecord(10) & "',0)"

Else
    'Existing Show.. update
     ShowID = CStr(rsInfo(0))
     db.Execute "UPDATE SHOWS set ShowName='" & rsRecord(1) & "',ShowStartDate=#" & rsRecord(2) & "#,ShowEndDate=#" & rsRecord(3) & "#,EventCity='" & rsRecord(4) & "',EventType='" & rsRecord(5) & "',EventReferenceName='" & rsRecord(6) & "',SQLShowID=" & rsRecord(9) & "  where ShowID=" & ShowID
End If
For i = 2 To 4
If rsInfo(i) Then
    Select Case i
    Case 2
     'delete ShowID record from SHOW_DEFAULTS
        If rsInfo(0) > 0 Then
            db.Execute "Delete * from SHOW_DEFAULTS where ShowID=" & ShowID
        End If
      'Get updated SHOW_DEFAULTS
        Set rsRecord = db2.OpenRecordset("Select * from SHOW_DEFAULTS where ShowID=" & ShowID)
      'Put into operational DB
        db.Execute "INSERT INTO SHOW_DEFAULTS values(" & rsRecord(0) & "," & rsRecord(1) & "," & rsRecord(2) & "," & rsRecord(3) & "," & rsRecord(4) & ")"
    Case 3
     'delete ShowID record from SHOW_TICKET_PRICES
        If rsInfo(0) > 0 Then
            db.Execute "Delete * from SHOW_TICKET_PRICES where ShowID=" & ShowID
        End If
      'Get updated SHOW_DEFAULTS
        Set rsRecord = db2.OpenRecordset("Select * from SHOW_TICKET_PRICES where ShowID=" & ShowID)
      'Put into operational DB
        Do Until rsRecord.EOF
            db.Execute "INSERT INTO SHOW_TICKET_PRICES values(" & rsRecord(0) & ",#" & rsRecord(1) & "#," & rsRecord(2) & "," & rsRecord(3) & "," & rsRecord(4) & ",'" & rsRecord(5) & "',#" & rsRecord(6) & "#)"
            rsRecord.MoveNext
        Loop
    Case 4
     'delete ShowID record from TICKET_HEADERS
        If rsInfo(0) > 0 Then
            db.Execute "Delete * from TICKET_HEADERS where ShowID=" & ShowID
        End If
      'Get updated SHOW_DEFAULTS
        Set rsRecord = db2.OpenRecordset("Select * from TICKET_HEADERS where ShowID=" & ShowID)
      'Put into operational DB
        db.Execute "INSERT INTO TICKET_HEADERS values(" & rsRecord(0) & ",'" & rsRecord(1) & "','" & rsRecord(2) & "','" & rsRecord(3) & "','" & rsRecord(4) & "','" & rsRecord(5) & "','" & rsRecord(6) & "','" & rsRecord(7) & "')"
    End Select
  End If
Next

Set rsInfo = Nothing
db.Close
db2.Close
wrkjet.Close
    MsgBox "DB Update Complete"
Else
    MsgBox "Error while downloading"
End If
End Sub

Sub UpdateEXE()
Dim errcode As Long
Dim url As String
Dim localFileName As String

url = "http://207.171.1.163/TicketmationUpdate.exe"
localFileName = App.Path & "\Ticketmation" & Format(Now, "mmddyy") & ".exe"

errcode = URLDownloadToFile(0, url, localFileName, 0, 0)
If errcode = 0 Then
    MsgBox "Download ok"
    Name App.Path & "\Ticketmation.exe" As App.Path & "\\TicketmationOld" & Format(Now, "mmddyy") & ".exe"
    FileCopy App.Path & "\Ticketmation" & Format(Now, "mmddyy") & ".exe", App.Path & "\Ticketmation.exe"
    MsgBox "EXE Update Complete"
Else
    MsgBox "Error while downloading"
End If


End Sub

Function UpdateRecord(strSQL)
Dim m_oConnection As ADODB.Connection
Dim oCommand As ADODB.Command
Dim oRS      As ADODB.Recordset
Set m_oConnection = New ADODB.Connection
m_oConnection.ConnectionString = strConnectionString
m_oConnection.CommandTimeout = 100000
m_oConnection.ConnectionTimeout = 100000
If bDebug Then MsgBox "Opening Database connection"
m_oConnection.Open
'// Open recordset
If bDebug Then MsgBox "Database connection Opened"

Set oCommand = New ADODB.Command
If bDebug Then MsgBox "Establishing Database connection "

oCommand.ActiveConnection = m_oConnection
If bDebug Then MsgBox "Database connection Established"

oCommand.CommandText = strSQL
oCommand.CommandType = adCmdText
Set oRS = oCommand.Execute()
If bDebug Then MsgBox "Database Returned"

'// Destroy object references
Set oCommand = Nothing

Set oRS = Nothing

End Function
 

Function GetRecordSet(strSQL)
Dim m_oConnection As ADODB.Connection
Dim oCommand As ADODB.Command
Dim oRS      As ADODB.Recordset
Set m_oConnection = New ADODB.Connection
m_oConnection.ConnectionString = strConnectionString
m_oConnection.CommandTimeout = 100000
m_oConnection.ConnectionTimeout = 100000
If bDebug Then MsgBox "Opening Database connection"
m_oConnection.Open
'// Open recordset
If bDebug Then MsgBox "Database connection Opened"

Set oCommand = New ADODB.Command
If bDebug Then MsgBox "Establishing Database connection "

oCommand.ActiveConnection = m_oConnection
If bDebug Then MsgBox "Database connection Established"

oCommand.CommandText = strSQL
oCommand.CommandType = adCmdText
Set oRS = oCommand.Execute()
If bDebug Then MsgBox "Database Returned"

'// Destroy object references
Set oCommand = Nothing

Set GetRecordSet = oRS
End Function

Function ProcIDFromWnd(ByVal hwnd As Long) As Long
   Dim idProc As Long
   
   ' Get PID for this HWnd
   GetWindowThreadProcessId hwnd, idProc
   
   ' Return PID
   ProcIDFromWnd = idProc
End Function
      
Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
   
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)
   
   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If
   
      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop
End Function



