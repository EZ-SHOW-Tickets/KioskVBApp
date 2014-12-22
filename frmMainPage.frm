VERSION 5.00
Begin VB.Form frmMainPage 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Background Ticketmation"
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMainPage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerCPUMonitor 
      Left            =   360
      Top             =   2520
   End
   Begin VB.Timer MonitorTimer 
      Left            =   360
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   1920
   End
   Begin VB.Image imAdmin 
      Height          =   1335
      Left            =   0
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Image ImDB 
      Height          =   1335
      Left            =   10440
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Image imUpdateApp 
      Height          =   1335
      Left            =   11520
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image ImageMonitor 
      Height          =   1215
      Left            =   8400
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblNotResponding 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "**************************"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   5160
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "THIS KIOSK IS CURRENTLY OUT OF SERVICE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   6975
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   12135
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bAllowed As Boolean
Dim bNoResponse As Boolean
Dim bDebug As Boolean
Dim hInst As Long             ' Instance handle from Shell function.


Private Sub Form_Load()
    Dim rsInfo As Recordset
'    On Error GoTo DB_Error
    Set wrkjet = CreateWorkspace("", "admin", "", dbUseJet)
    Dim strpath As String
    strpath = App.Path & "\WUS.mdb"
    Set db = wrkjet.OpenDatabase(strpath)
    Set rsInfo = db.OpenRecordset("Select IPAddress from INIT")
IP = Str(rsInfo(0))
Set rsInfo = Nothing
db.Close
wrkjet.Close
bAllowed = True
bNoResponse = False
bDebug = False
'TimerCPUMonitor.Interval = 5000
'TimerCPUMonitor.Enabled = True
'MonitorTimer.Interval = 10000
'MonitorTimer.Enabled = True
End Sub

Private Sub imAdmin_DblClick()
fAdmin.Show 1
End Sub


Private Sub Image1_DblClick()
End
End Sub


Private Sub Image2_DblClick()
Dim Response As Integer
'Start App
If bAllowed Then
    Response = MsgBox("Run TICKETMATION with no monitor?", vbYesNo, "RUN")
    If Response = 7 Then
        Exit Sub
    Else
        ' Shell to an application
        '   hInst = Shell("C:\CGC\Projects\Ticketmation\NEW_VB_APP\Ticketmation.exe", vbMaximizedFocus)
        hInst = Shell(App.Path & "\Ticketmation.exe", vbMaximizedFocus)
        'hInst = Shell(App.Path & "\NonResponse.exe", vbMaximizedFocus)
    '    If bDebug Then MsgBox "hInst=" & Str(hInst)
    '
    ''    ' Get Window handle
    ''    hWndApp = GetWinHandle(hInst)
    '    'Begin Monitor
    ''    If bDebug Then MsgBox "hWndApp=" & Str(hWndApp)
    '
    '    MonitorTimer.Interval = 30000
    '    MonitorTimer.Enabled = True
    '    LblNotResponding.Visible = False
    
        Timer1.Interval = 5000
        Timer1.Enabled = True
        bAllowed = False
    End If
End If

End Sub



Private Sub ImageMonitor_DblClick()
'Start App
If bAllowed Then
    Response = MsgBox("Run TICKETMATION with monitor?", vbYesNo, "RUN")
    If Response = 7 Then
        Exit Sub
    Else
    ' Shell to an application
    '   hInst = Shell("C:\CGC\Projects\Ticketmation\NEW_VB_APP\Ticketmation.exe", vbMaximizedFocus)
    hInst = Shell(App.Path & "\Ticketmation.exe", vbMaximizedFocus)
    'hInst = Shell(App.Path & "\NonResponse.exe", vbMaximizedFocus)
    If bDebug Then MsgBox "hInst=" & Str(hInst)
    
'    ' Get Window handle
'    hWndApp = GetWinHandle(hInst)
    'Begin Monitor
'    If bDebug Then MsgBox "hWndApp=" & Str(hWndApp)

    MonitorTimer.Interval = 30000
    MonitorTimer.Enabled = True
    LblNotResponding.Visible = False

    Timer1.Interval = 5000
    Timer1.Enabled = True
    bAllowed = False
  End If
End If

End Sub


Private Sub ImDB_Click()
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

Private Sub imUpdateApp_DblClick()

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


Private Sub imUpdateSuper_Click()

End Sub


Private Sub imUpdateSuper_DblClick()
Dim errcode As Long
Dim url As String
Dim localFileName As String

url = "http://207.171.1.163/SuperAppUpdate.exe"
localFileName = "c:\Ticketmation\SuperApp" & Format(Now, "mmddyy") & ".exe"

errcode = URLDownloadToFile(0, url, localFileName, 0, 0)
If errcode = 0 Then
    MsgBox "Download ok"
    Name "c:\Ticketmation\SuperApp.exe" As "c:\Ticketmation\SuperAppOld" & Format(Now, "mmddyy") & ".exe"
    FileCopy "c:\Ticketmation\SuperApp" & Format(Now, "mmddyy") & ".exe", "c:\Ticketmation\TSuperApp.exe"
    MsgBox "EXE Update Complete"
Else
    MsgBox "Error while downloading"
End If
End Sub


Private Sub MonitorTimer_Timer()
'Dim sappname As String

'sappname = "C:\CGC\Projects\Ticketmation\NEW_VB_APP\ticketmation.exe"
''sAppPath = "D:\program files\Adobe\Photoshop 6.0\Photoshp.exe"
''If App.PrevInstance Then
'If IsTaskRunning(sappname) Then
'    MsgBox "Application '" & sappname & "' is running!"
'End If
' Get Window handle
hWndApp = GetWinHandle(hInst)
lngReturnValue = SendMessageTimeout(hWndApp, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, 1000, lngResult)
If bDebug Then MsgBox "hWndApp=" & Str(hWndApp) & " lngReturnValue=" & Str(lngReturnValue)
If lngReturnValue Then
    'MsgBox "Responding"
'    bNoResponse = False
'    MonitorTimer.Enabled = True
    LblNotResponding.Visible = False
ElseIf bNoResponse Then
    'MsgBox "Kill it"
    'Kill IT
    lngReturnValue = GetWindowThreadProcessId(hWndApp, lngProcessID)
    lngProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lngProcessID)
    lngReturnValue = TerminateProcess(lngProcess, 0&)
    MonitorTimer.Enabled = False
    LblNotResponding.Visible = True
    UpdateRecord ("UPDATE DEVICE_ACTIVE set DeviceActive='X' where DeviceID=" & IP)
Else
    'MsgBox "Not Responding"
    bNoResponse = True
End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    Timer1.Enabled = False
    bAllowed = True
    UpdateRecord ("UPDATE DEVICE_ACTIVE set DeviceActive='A' where DeviceID=" & IP)
    End Sub


Private Sub TimerCPUMonitor_Timer()
MsgBox CPU_Usage_Percent
MsgBox "Memory Load=" & GetMemoryLoad
End Sub


