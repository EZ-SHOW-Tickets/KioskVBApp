VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmWebBrowser 
   BorderStyle     =   0  'None
   Caption         =   "Please Wait, PROCESSING..."
   ClientHeight    =   90
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timerwait 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Left            =   30
      Top             =   750
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   135
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5145
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   10050
      ExtentX         =   17727
      ExtentY         =   9075
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim strTixSold As String
Dim iTixSold As Integer
Dim iBatch As String
Dim i

If bDebug Then MsgBox "WebBrowser Load"
If bBatch Then
    iBatch = "1"
Else
    iBatch = "0"
End If
strTixSold = ""
iTixSold = 0
For i = 0 To UBound(TicketTypes, 2)
    strTixSold = strTixSold & "&TType" & Trim(str(i)) & "=" & TicketTypes(1, i) & "&TTypeN" & Trim(str(i)) & "=" & TotalTicketSold(i)
    iTixSold = iTixSold + TotalTicketSold(i)
Next i
'webCounter = 0
'If bDebug Then MsgBox "SEND to AN ?Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Track1a=" & Track1 & "&Device=" & IP & "&Name=" & CardName & "&Stock=" & TicketStock - iTixSold & strTixSold
'If Len(Trim(ZipCode)) <> 5 Then
'    ZipCode = "99999"
'End If
'WebBrowser1.Navigate "https://www.cgc-services.com/ticketmation/kiosk/ProcessPaymentfromkiosk.asp?Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&STime=" & SwipTime & "&ZipCode=" & ZipCode & "&SQLShowid=" & SQLShowID & "&Coupon=" & CouponCode & "&Stock=" & TicketStock - iTixSold & strTixSold
'Removed Coupon and ZipCode
'WebBrowser1.Navigate "https://www.cgc-services.com/ticketmation/kiosk/ProcessPaymentfromkiosk.asp?Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&STime=" & SwipTime & "&SQLShowid=" & SQLShowID & "&Stock=" & TicketStock - iTixSold & strTixSold
'Added Batch Indicator
'USE THIS

'If bUseBU Then
'    If Len(Trim(ZipCode)) <> 5 Then
'        WebBrowser1.Navigate "http://54.241.244.28/ticketmation/kiosk/ProcessPaymentfromkioskV3.asp?BCode=" & Format(IP, "0#") & Trim(SessionCounter) & Trim(str(SessionID)) & "&Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&STime=" & SwipTime & "&SQLShowid=" & SQLShowID & "&Batch=" & iBatch & "&Stock=" & TicketStock - iTixSold & strTixSold
'    Else
'        WebBrowser1.Navigate "http://54.241.244.28/ticketmation/kiosk/ProcessPaymentfromkioskV3.asp?BCode=" & Format(IP, "0#") & Trim(SessionCounter) & Trim(str(SessionID)) & "&Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&ZIP=" & ZipCode & "&STime=" & SwipTime & "&SQLShowid=" & SQLShowID & "&Batch=" & iBatch & "&Stock=" & TicketStock - iTixSold & strTixSold
'    End If
'Else
    Timerwait.interval = 10000 'Wait 10 seconds for web response
    Timerwait.Enabled = True
    If Len(Trim(ZipCode)) <> 5 Then
        WebBrowser1.Navigate "https://www.cgc-services.com/ticketmation/kiosk/ProcessPaymentfromkioskV3.asp?BCode=" & Format(IP, "0#") & Trim(SessionCounter) & Trim(str(SessionID)) & "&Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&STime=" & SwipTime & "&SQLShowid=" & SQLShowID & "&Batch=" & iBatch & "&Stock=" & TicketStock - iTixSold & strTixSold
    Else
        WebBrowser1.Navigate "https://www.cgc-services.com/ticketmation/kiosk/ProcessPaymentfromkioskV3.asp?BCode=" & Format(IP, "0#") & Trim(SessionCounter) & Trim(str(SessionID)) & "&Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&ZIP=" & ZipCode & "&STime=" & SwipTime & "&SQLShowid=" & SQLShowID & "&Batch=" & iBatch & "&Stock=" & TicketStock - iTixSold & strTixSold
    End If
'End If
'TEST BATCH
'WebBrowser1.Navigate "https://www.cgc-servicesxx.com/eztv5/ProcessPaymentfromkiosk.asp?Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&STime=" & SwipTime & "&Stock=" & TicketStock - iTixSold & strTixSold
'Debug.Print "https://www.cgc-servicesxx.com/eztv5/ProcessPaymentfromkiosk.asp?Amount=" & Format(TotalPrice, "0#.00") & "&Track1=" & AccountNumber & "&Device=" & IP & "&Name=" & CardName & "&STime=" & SwipTime & "&ZipCode=" & ZipCode & "&Stock=" & TicketStock - iTixSold & strTixSold
'BackupServer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WebBrowser1.Navigate "https://www.cgc-services.com/eztv5/ProcessPaymentFromKiosk.asp?Amount=" & TotalPrice & "&Track1=" & AccountNumber
'    If WebBrowser1.Busy Then
'        MsgBox "Can't access the document while a download operation is in progress."
'        Exit Sub
'    End If
'    If Inet1.StillExecuting Then
'        MsgBox "Internet Transfer control is busy downloading at the time."
'        Exit Sub
'    End If
'
'    Dim msg As String
'    Dim i As Integer
'    Dim j As Integer
'    Dim x As Integer
'    'msg = "The page contains " & WebBrowser1.Document.Forms.Length & " FORM objects" & vbCrLf
'    For i = 0 To WebBrowser1.Document.Forms.Length - 1
'        'msg = msg & "   FORM NAME = " & WebBrowser1.Document.Forms(i).Name & vbCrLf
'        'msg = msg & "      ACTION = " & WebBrowser1.Document.Forms(i).Action & vbCrLf
'        For j = 0 To WebBrowser1.Document.Forms(i).Length - 1
'             Select Case WebBrowser1.Document.Forms(i).Elements(j).Name
'                Case "ANResponse"
'                    ANResponseCode = CInt(WebBrowser1.Document.Forms(i).Elements(j).Value)
'                Case "ANTransactionNumber"
'                    ANTransactionCode = WebBrowser1.Document.Forms(i).Elements(j).Value
'             End Select
'           msg = msg & "   ELEMENT [" & WebBrowser1.Document.Forms(i).Elements(j).Name & "]};" & vbCrLf
'            msg = msg & "   has a value of   [" & WebBrowser1.Document.Forms(i).Elements(j).Value & "]" & vbCrLf
'        Next j
'    Next i
''    Screen.MousePointer = vbDefault
''    Debug.Print msg
'Timer2.Interval = 100
'Timer2.Enabled = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub


Private Sub Timer2_Timer()
webCounter = webCounter + 1
    If WebBrowser1.Busy Then
            Timer2.interval = 0
            Timer2.Enabled = False
            Dim msg As String
            Dim i As Integer
            Dim j As Integer
            Dim x As Integer
            'msg = "The page contains " & WebBrowser1.Document.Forms.Length & " FORM objects" & vbCrLf
            For i = 1 To WebBrowser1.Document.Forms.length - 1
                'msg = msg & "   FORM NAME = " & WebBrowser1.Document.Forms(i).Name & vbCrLf
                'msg = msg & "      ACTION = " & WebBrowser1.Document.Forms(i).Action & vbCrLf
                For j = 0 To WebBrowser1.Document.Forms(i).length - 1

                     Select Case WebBrowser1.Document.Forms(i).Elements(j).Name
                        Case "ANResponse"
                            ANResponseCode = WebBrowser1.Document.Forms(i).Elements(j).Value
                        Case "ANTransactionNumber"
                            ANTransactionCode = WebBrowser1.Document.Forms(i).Elements(j).Value
                     End Select
                     msg = msg & "   ELEMENT [" & WebBrowser1.Document.Forms(i).Elements(j).Name & "]};" & vbCrLf
                    msg = msg & "   has a value of   [" & WebBrowser1.Document.Forms(i).Elements(j).Value & "]" & vbCrLf
                Next j
            Next i
            Debug.Print msg

            Unload Me
    Else
        Timer2.interval = 100
    End If
End Sub


Private Sub Timer3_Timer()

End Sub

Private Sub Timerwait_Timer()
WebBrowser1.Stop
Dim strSQL As String
Timerwait.Enabled = False
If Not bBatch Then
    If bDebug Then MsgBox "BATCH MODE"
    
    db.Execute "INSERT INTO BATCH (track1,track2,TotalPrice,BatchDateTime,ZipCode) values('" & Track1 & "','" & AccountNumber & "'," & TotalPrice & ",#" & Now() & "#,'" & ZipCode & "')"
    Set rs = db.OpenRecordset("Select Max(BatchID) as TRID from BATCH")
    'nBatch = nBatch + 1
    'Set rs = db.OpenRecordset("SELECT max ")
    ''Set rs = conn.Execute("INSERT into someTable(IntColumn) values (" & fakeValue & "); SELECT @@IDENTITY as newID").NextRecordset
    'strSQL = "INSERT INTO BATCH (track1,track2,TotalPrice,BatchDateTime) values('" & Track1 & "','" & AccountNumber & "'," & TotalPrice & ",#" & Now() & "#);SELECT @@IDENTITY"
    
    'Set rs = db.Execute(strSQL)
    
    For i = 0 To UBound(TicketTypes, 2)
        If TotalTicketSold(i) > 0 Then db.Execute "Insert into BATCH_DETAILS (BatchID,TicketTypeID,NumberTix) values(" & rs(0) & "," & TicketTypes(1, i) & "," & TotalTicketSold(i) & ")"
    Next i
    Call LogClick("BatchMode", str(rs(0)))
    
    ANResponseCode = 1
    Randomize Timer
    randomnumber = Int((99999999 - 1234567 + 1) * Rnd) + 1234567
    ANTransactionCode = "-1" & Format(rs(0), "00#") & Right(str(randomnumber), 6)
End If
Unload Me
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)
            Dim msg As String
            Dim i As Integer
            Dim j As Integer
            Dim x As Integer
            Dim ANResponseText
            'msg = "The page contains " & WebBrowser1.Document.Forms.Length & " FORM objects" & vbCrLf
            If bDebug Then MsgBox "Back from AN.  Length=" & WebBrowser1.Document.Forms.length

            For i = 0 To WebBrowser1.Document.Forms.length - 1
                'msg = msg & "   FORM NAME = " & WebBrowser1.Document.Forms(i).Name & vbCrLf
                'msg = msg & "      ACTION = " & WebBrowser1.Document.Forms(i).Action & vbCrLf
                For j = 0 To WebBrowser1.Document.Forms(i).length - 1

                     Select Case WebBrowser1.Document.Forms(i).Elements(j).Name
                        Case "ANResponse"
                            ANResponseCode = WebBrowser1.Document.Forms(i).Elements(j).Value
                        Case "ANTransactionCode"
                            ANTransactionCode = WebBrowser1.Document.Forms(i).Elements(j).Value
                        Case "ANResponseTxt"
                            ANResponseText = WebBrowser1.Document.Forms(i).Elements(j).Value
                     End Select
                    msg = msg & "   ELEMENT [" & WebBrowser1.Document.Forms(i).Elements(j).Name & "]};" & vbCrLf
                    msg = msg & "   has a value of   [" & WebBrowser1.Document.Forms(i).Elements(j).Value & "]" & vbCrLf
                Next j
            Next i
           ' Debug.Print msg
            If bDebug Then MsgBox msg
            Unload Me

End Sub

Private Sub WebBrowser1_NavigateError(ByVal pDisp As Object, url As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
'Do not process if in batch upload
Dim strSQL As String

If Not bBatch Then
    If bDebug Then MsgBox "BATCH MODE"
    
    db.Execute "INSERT INTO BATCH (track1,track2,TotalPrice,BatchDateTime,ZipCode) values('" & Track1 & "','" & AccountNumber & "'," & TotalPrice & ",#" & Now() & "#,'" & ZipCode & "')"
    Set rs = db.OpenRecordset("Select Max(BatchID) as TRID from BATCH")
    'nBatch = nBatch + 1
    'Set rs = db.OpenRecordset("SELECT max ")
    ''Set rs = conn.Execute("INSERT into someTable(IntColumn) values (" & fakeValue & "); SELECT @@IDENTITY as newID").NextRecordset
    'strSQL = "INSERT INTO BATCH (track1,track2,TotalPrice,BatchDateTime) values('" & Track1 & "','" & AccountNumber & "'," & TotalPrice & ",#" & Now() & "#);SELECT @@IDENTITY"
    
    'Set rs = db.Execute(strSQL)
    
    For i = 0 To UBound(TicketTypes, 2)
        If TotalTicketSold(i) > 0 Then db.Execute "Insert into BATCH_DETAILS (BatchID,TicketTypeID,NumberTix) values(" & rs(0) & "," & TicketTypes(1, i) & "," & TotalTicketSold(i) & ")"
    Next i
    Call LogClick("BatchMode", str(rs(0)))
    
    ANResponseCode = 1
    Randomize Timer
    randomnumber = Int((99999999 - 1234567 + 1) * Rnd) + 1234567
    ANTransactionCode = "-1" & Format(rs(0), "00#") & Right(str(randomnumber), 6)
End If
Unload Me

End Sub


