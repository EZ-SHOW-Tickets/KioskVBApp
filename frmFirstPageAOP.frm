VERSION 5.00
Object = "{158336E7-3FF3-456E-912C-5985E9BBED24}#1.2#0"; "MTUSBHIDSwipe.ocx"
Begin VB.Form frmFirstPageAOP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Ticketmation"
   ClientHeight    =   11475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11475
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerBadSwipe 
      Left            =   13920
      Top             =   3120
   End
   Begin VB.Timer Timer_idleMinutes 
      Left            =   13890
      Top             =   2325
   End
   Begin VB.Timer TimerWrongCard 
      Left            =   13875
      Top             =   1860
   End
   Begin VB.Timer TimerStupid 
      Left            =   13845
      Top             =   1320
   End
   Begin VB.Timer TimerClick 
      Left            =   13830
      Top             =   810
   End
   Begin VB.Timer ResetTimer 
      Left            =   13845
      Top             =   285
   End
   Begin ctlUSBHID.USBHID USBHID1 
      Left            =   12825
      Top             =   330
      _ExtentX        =   1349
      _ExtentY        =   1482
   End
   Begin VB.Label lblNoDiscounts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "THIS MACHINE DOES NOT ACCEPT DISCOUNTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   4080
      Width           =   9975
   End
   Begin VB.Label lOffLine 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   330
      Left            =   495
      TabIndex        =   3
      Top             =   10725
      Width           =   480
   End
   Begin VB.Image imTest 
      Height          =   615
      Left            =   960
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Image ImBadSwipe 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Left            =   5400
      Picture         =   "frmFirstPageAOP.frx":0000
      Top             =   4440
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.Image imgAOP_Logo 
      Height          =   1335
      Left            =   240
      Picture         =   "frmFirstPageAOP.frx":67B5
      Top             =   120
      Width           =   2880
   End
   Begin VB.Label lblBadSwipe 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "THIS MACHINE DOES NOT ACCEPT THIS TYPE OF CARD   PLEASE USE A DIFFERENT CARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   840
      TabIndex        =   10
      Top             =   8640
      Visible         =   0   'False
      Width           =   13215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SWIPE CREDIT CARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   2528
      TabIndex        =   9
      Top             =   5400
      Width           =   9735
   End
   Begin VB.Label lblOr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "or"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5408
      TabIndex        =   8
      Top             =   6600
      Width           =   3975
   End
   Begin VB.Label lblTouch 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Touch Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   3488
      TabIndex        =   7
      Top             =   7320
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   3668
      TabIndex        =   6
      Top             =   4680
      Width           =   7455
   End
   Begin VB.Image ImgSwipe 
      Height          =   3090
      Left            =   840
      Picture         =   "frmFirstPageAOP.frx":8AA3
      Top             =   5040
      Width           =   2220
   End
   Begin VB.Image ImgFinger 
      Height          =   1875
      Left            =   10920
      Picture         =   "frmFirstPageAOP.frx":B937
      Top             =   6960
      Width           =   2160
   End
   Begin VB.Image imgCC 
      Height          =   300
      Index           =   3
      Left            =   9840
      Picture         =   "frmFirstPageAOP.frx":CB76
      Top             =   10080
      Width           =   1665
   End
   Begin VB.Image imgCC 
      Height          =   915
      Index           =   2
      Left            =   7575
      Picture         =   "frmFirstPageAOP.frx":D3C3
      Top             =   9780
      Width           =   1545
   End
   Begin VB.Image imgCC 
      Height          =   810
      Index           =   1
      Left            =   5595
      Picture         =   "frmFirstPageAOP.frx":E741
      Top             =   9825
      Width           =   1275
   End
   Begin VB.Image imgCC 
      Height          =   675
      Index           =   0
      Left            =   4200
      Picture         =   "frmFirstPageAOP.frx":FFD2
      Top             =   9900
      Width           =   675
   End
   Begin VB.Label lblPurchaseHere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE TICKETS HERE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   2115
      TabIndex        =   5
      Top             =   3240
      Width           =   10575
   End
   Begin VB.Image ImgShow 
      Height          =   3000
      Left            =   0
      Top             =   0
      Width           =   15360
   End
   Begin VB.Label LabelSwipeClear 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   13080
      TabIndex        =   4
      Top             =   7200
      Width           =   1200
   End
   Begin VB.Label TXTbought 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   11130
      Width           =   585
   End
   Begin VB.Label ClickLabel2 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   525
      Left            =   13905
      TabIndex        =   1
      Top             =   10050
      Width           =   465
   End
   Begin VB.Label ClickLabel1 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   525
      Left            =   465
      TabIndex        =   0
      Top             =   10125
      Width           =   465
   End
End
Attribute VB_Name = "frmFirstPageAOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bTouch As Boolean
Dim ClickClick1 As Boolean
Dim idleMinutes As Integer
Dim WhichScreen As Integer

Sub PageClicked()
If bTouch Then
    bTouch = False
    ResetTimer.interval = 10000
    ResetTimer.Enabled = True
    Me.BackColor = &HC0C0FF
    ImgFinger.Visible = False
    lblOr.Visible = False
    lblTouch.Visible = False
'    Image1.Picture = LoadPicture(App.Path & "/SecondCCPage" & strPageImageCC & ".jpg")
End If


End Sub


Private Sub ClickLabel1_Click()
ClickClick1 = True
TimerClick.interval = 5000
TimerClick.Enabled = True

End Sub

Private Sub ClickLabel2_Click()
If ClickClick1 Then
        TimerClick.interval = 0
        TimerClick.Enabled = False
        ClickClick1 = False
        fTicketmation.Show 1
        
End If

End Sub


Private Sub Form_Click()
PageClicked
End Sub

Private Sub Form_Load()
Dim rsBatch As Recordset
Dim rsTicketstock As Recordset
''Load appropriate image
'''NO LOAD FOR DUB
On Error Resume Next
If ShowID = 59 Then
    ImgShow.Top = 0
    ImgShow.Picture = LoadPicture(App.path & "/" & ShowImage)
    'ImgShow.Left = 7395 - ImgShow.Width / 2
    ImgShow.Left = 0
    imgAOP_Logo.Visible = False
Else
    ImgShow.Top = 1320
    ImgShow.Picture = LoadPicture(App.path & "/" & ShowImage)
    ImgShow.Left = 7395 - ImgShow.Width / 2
    imgAOP_Logo.Visible = True
End If
If AOPBarcode Then
    imgAOP_Logo.Picture = LoadPicture(App.path & "/AOP_logo.gif")
Else
    imgAOP_Logo.Picture = LoadPicture(App.path & "/ticketmation_logo_trans4.gif")
End If

iSwipe = 0
Me.BackColor = &HFFC0C0
'specify credit card icons
If VisaMCOnly Then
    imgCC(0).Visible = False
Else
    imgCC(0).Visible = True
End If
If AllowDiscover Then
    imgCC(3).Visible = True
Else
    imgCC(3).Visible = False
  
End If
PrinterError = False
'Open Card Read Port
USBHID1.PortOpen = True
If Not (USBHID1.PortOpen) Then
''    MsgBox "Couldn't open HID Swipe Reader"
    Call LogClick("Load", "HID Not Open")
    'End
End If
'set reset swipe timer
idleMinutes = 0
Timer_idleMinutes.interval = 60000
Timer_idleMinutes.Enabled = True
'Set screen to allow start on TOUCH
bTouch = True
'Turn off Timer for Screen2
ResetTimer.interval = 0
ResetTimer.Enabled = False
'Check if DB indicates BATCH data
Set rsBatch = db.OpenRecordset("Select count(*) as noBatches from BATCH where BatchUpload=No")
'Set Batch Indicator on screen
'If rsBatch.EOF Then
If rsBatch(0) = 0 Then
    nBatch = 0
    'lOffLine.Caption = "0"
Else
    nBatch = rsBatch(0)
    'lOffLine.Caption = "*"
End If
lOffLine.Caption = Trim(str(nBatch))

'Check Ticket stock
Set rsTicketstock = db.OpenRecordset("Select * from TICKET_STOCK")
If rsTicketstock(0) <= 10 Then
    frmStartup.Show
    Unload Me
Else
    TXTbought.Caption = rsTicketstock(0)
    TicketStock = rsTicketstock(0)
End If
SwipeProcess = True

End Sub






Private Sub imgAOP_Logo_Click()
PageClicked
End Sub

Private Sub imgCC_Click(Index As Integer)
PageClicked
End Sub


Private Sub ImgFinger_Click()
PageClicked
End Sub

Private Sub ImgShow_Click()
PageClicked
End Sub


Private Sub ImgSwipe_Click()
PageClicked
End Sub

Private Sub imTest_Click()
If bDebug Then
    Dim i As Integer
    i = 0
    Do Until i > 1
        i = 0
    Loop
End If
End Sub

Private Sub Label1_Click()
PageClicked
End Sub



Private Sub Label4_Click()
PageClicked
End Sub


Private Sub LabelSwipeClear_DblClick()
Dim lresult As Long
Dim output As String
MsgBox "CLEAR"
'USBHID2.PortOpen = True
lresult = USBHID1.USBSwipe_Command("02", output)
'MsgBox ("Result=" & CStr(lresult) & ",Data=" & USBHID2.SwipeCommandOutput)
USBHID1.PortOpen = True


End Sub


Private Sub lblBadSwipe_Click()
PageClicked
End Sub

Private Sub lblOr_Click()
PageClicked
End Sub

Private Sub lblPurchaseHere_Click()
PageClicked
End Sub

Private Sub lblTouch_Click()
PageClicked
End Sub

Private Sub ResetTimer_Timer()
ResetTimer.interval = 0
ResetTimer.Enabled = False
bTouch = True
'Image1.Picture = LoadPicture(App.Path & "/FirstPageCC" & strPageImageCC & ".jpg")
    Me.BackColor = &HFFC0C0
    ImgFinger.Visible = True
    lblOr.Visible = True
    lblTouch.Visible = True


End Sub


Private Sub Timer_idleMinutes_Timer()
Dim lresult As Long
Dim output As String

idleMinutes = idleMinutes + 1
'If idleMinutes > 5 And Len(output) > 0 Then
If idleMinutes > 5 Then
    lresult = USBHID1.USBSwipe_Command("02", output)
    USBHID1.PortOpen = True
    idleMinutes = 0
End If

Timer_idleMinutes.interval = 60000

End Sub

Private Sub TimerBadSwipe_Timer()
'TimerBadSwipe.Interval = 0
TimerBadSwipe.Enabled = False

ImBadSwipe.Visible = False
End Sub

Private Sub TimerClick_Timer()
ClickClick1 = False
'TimerClick.Interval = 0
TimerClick.Enabled = False
End Sub


Private Sub TimerStupid_Timer()
            Dim x As Integer
            'onerror
            Call LogClick("TimerStupid", "Enter")
            
            On Error GoTo SWipeError

            If bDebug Then MsgBox "In STUPID"
            'TimerStupid.Interval = 0
            TimerStupid.Enabled = False
            'MsgBox "OPENING frmSpecifyTickets"
            'x = 1
            'frmKluge.Show 1
             Call LogClick("TimerStupid", "Ready For Specify")
             frmSpecifyTickets.Show
             Call LogClick("TimerStupid", "Ready For Unload")
             'USBHID1.PortOpen = False
             SwipeProcess = True
             Unload Me
             Exit Sub
SWipeError:

          MsgBox "ERROR" & Err.Description & " " & Err.Number
          End
End Sub

Private Sub TimerWrongCard_Timer()
'lblBadSwipe.Caption = ""
lblBadSwipe.Visible = False
'TimerWrongCard.Interval = 0
TimerWrongCard.Enabled = False

End Sub

Private Sub USBHID1_CardDataChanged()
If SwipeProcess Then
'    SwipeProcess = False
    Dim abTrack1() As Byte
    Dim abTrack2() As Byte
    Dim abTrack3() As Byte
    Dim i As Long
    Dim j As Integer
    Dim strOutput As String
    ''Text1.Text = ""
    'Unload frmBadSwipe
    On Error GoTo SWipeError
    ImBadSwipe.Visible = False
    lblBadSwipe.Visible = False
    TimerBadSwipe.interval = 0
    TimerBadSwipe.Enabled = False
    Call LogClick("Swipe", "Process Swipe")
    
    iSwipe = 0
    Timer_idleMinutes.interval = 0
    Timer_idleMinutes.Enabled = False
    With USBHID1
        If bDebug Then
            MsgBox "Track1:" & .GetTrack(1)
            MsgBox "Track2:" & .GetTrack(2)
            MsgBox "Track3:" & .GetTrack(3)
        End If
        If (.GetTrack(1) <> "" Or .GetTrack(2) <> "" Or .GetTrack(3) <> "") Then
               If InStr(1, .GetTrack(2), "=") = 0 Then
                    'lblBadSwipe.Caption = "PLEASE USE A VALID CARD."
                    lblBadSwipe.Visible = True
                    Call LogClick("Swipe", "Invalid Card")
                    TimerWrongCard.interval = 3000
                    TimerWrongCard.Enabled = True
                    'SwipeProcess = True
                    Exit Sub
               Else
                    SwipTime = Now()
                    AccountNumber = .GetTrack(2)
                    If Len(AccountNumber) > 0 Then
                        CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
                        CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
                                'Credit Card Type
                        If Left(CardNumber, 2) = "34" Or Left(CardNumber, 2) = "37" Then
                            MOP = "AMX"
                        ElseIf Left(CardNumber, 1) = "5" Then
                            MOP = "MC"
                        ElseIf Left(CardNumber, 1) = "4" Then
                            MOP = "VISA"
                        ElseIf Left(CardNumber, 3) = "300" Or Left(CardNumber, 3) = "303" Or Left(CardNumber, 3) = "302" Or Left(CardNumber, 3) = "303" Or Left(CardNumber, 3) = "304" Or Left(CardNumber, 3) = "305" Or Left(CardNumber, 2) = "36" Or Left(CardNumber, 2) = "38" Then
                            MOP = "DSC"
                        ElseIf Left(CardNumber, 4) = "2131" Or Left(CardNumber, 4) = "1800" Then
                            MOP = "JCB"
                        ElseIf Left(CardNumber, 6) = "628181" Then
                            MOP = "SEARS"
                        Else
                            MOP = "OTHER"
                        End If
            
            ''
                        If bDebug Then MsgBox "Card Type:" & Left(CardNumber, 1) & "," & VisaMCOnly
                        If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
                            'Card is a debit
                            'lblWelcome(3).FontSize = 16
                            'lblWelcome(3).ForeColor = &HFF&
                            'lblBadSwipe.Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
                            lblBadSwipe.Visible = True
                            'Timer3.Interval = 3000
                            'Timer3.Enabled = True
                            Call LogClick("Swipe", "No Debit Accepted")
                            TimerWrongCard.interval = 3000
                            TimerWrongCard.Enabled = True
                            'SwipeProcess = True
                            Exit Sub
                        'ElseIf VisaMCOnly And Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4") And (CardType <> "120" And CardType <> "220") Then
                        ElseIf VisaMCOnly And Left(CardNumber, 1) = "3" Then
                            'Card is an AMEX
                            'lblWelcome(3).FontSize = 16
                            'lblWelcome(3).ForeColor = &HFF&
                            If NotAllowDebit Then
                                'lblBadSwipe.Caption = "THIS MACHINE ONLY ACCEPTS VISA AND MASTERCARD."
                                lblBadSwipe.Visible = True
                            Else
                                'lblBadSwipe.Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD OR DEBIT CARDS."
                                lblBadSwipe.Visible = True
                            End If
                            Call LogClick("Swipe", "No AMEX Accepted")

                            TimerWrongCard.interval = 3000
                            TimerWrongCard.Enabled = True
                            
                            'Timer3.Interval = 3000
                            'Timer3.Enabled = True
                            'SwipeProcess = True
                            Exit Sub
                        ElseIf Left(CardNumber, 4) = "6011" And Not AllowDiscover Then
                            'Card is a debit
                            'lblWelcome(3).FontSize = 16
                            'lblWelcome(3).ForeColor = &HFF&
                            'lblBadSwipe.Caption = "THIS MACHINE DOES NOT ACCEPT DISCOVER CARDS.  PLEASE USE A DIFFERENT CARD."
                            lblBadSwipe.Visible = True
                            'Timer3.Interval = 3000
                            'Timer3.Enabled = True
                            Call LogClick("Swipe", "No Discover Accepted")
                            
                            TimerWrongCard.interval = 3000
                            TimerWrongCard.Enabled = True
                            'SwipeProcess = True
                            Exit Sub
                        ElseIf Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4" And Left(CardNumber, 1) <> "3" And Left(CardNumber, 4) <> "6011" And CardType <> "120" And CardType <> "220" Then
                            'lblWelcome(3).FontSize = 16
                            'lblWelcome(3).ForeColor = &HFF&
                            'lblBadSwipe.Caption = "PLEASE USE A VALID CARD."
                            lblBadSwipe.Visible = True
                            'Timer3.Interval = 3000
                            'Timer3.Enabled = True
                            Call LogClick("Swipe", "Invalid Card")
                            TimerWrongCard.interval = 3000
                            TimerWrongCard.Enabled = True
                            'SwipeProcess = True
                            Exit Sub
                        End If
                        ''NO PROBLEMS ... CARD OK - Get data
                        Timer_idleMinutes.interval = 0
                        Timer_idleMinutes.Enabled = False
            
                        CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
                        Track1 = .GetTrack(1)
                        If bDebug Then MsgBox "Track1=" & Track1
                        If bDebug Then MsgBox "Track2=" & AccountNumber
                        'For i = 1 To 1000000
                        '     DoEvents
                        'Next i
                        If Len(Track1) > 1 Then
                            CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
                        Else
                            CardName = .GetLName & "/" & .GetFName
                        End If
                        'Close Swipe Port
                        USBHID1.ClearBuffer
                        USBHID1.PortOpen = False
                        'Start New session
                        db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
                        Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
                        SessionID = rs(0)
                        Call LogClick("Swipe", "Card Accepted")
                        If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
                        If bDebug Then MsgBox "OPENING frmSpecifyTickets"
            ''''''''''''''''''USE TIMER STUPID
            '           TimerStupid.Interval = 600
            '           Call LogClick("Swipe", "Timer Set")
            '''''''''''''''''''''''''''''''''''
                    '' Turn all timers OFF
                        Timer_idleMinutes.Enabled = False
                        TimerBadSwipe.Enabled = False
                        TimerClick.Enabled = False
                        TimerStupid.Enabled = False
                        TimerWrongCard.Enabled = False
                    'For j = 1 To 2
 
                   '     me.Hide
        ''                Call LogClick("Swipe", "Show Specify")
        ''           '     frmSpecifyTickets.Show
        ''                TimerStupid.Interval = 400
        ''                TimerStupid.Enabled = True
        ''                Call LogClick("Swipe", "Timer Set")
        
                    '    Unload Me
            '            Next j
            '            MsgBox "OPENING frmSpecifyTickets"
                        'frmKluge.Show
                        'Unload Me
    '''''use kluge'''''''''''
    '''PUT THIS BACK
                        Call LogClick("After Swipe", "To KLUGE")
                        SwipeProcess = True
                        For i = 1 To 100
                             DoEvents
                        Next i
                        frmKluge.Show
'                        Call LogClick("After Swipe", "From KLUGE")
'                        frmSpecifyTickets.Show
    ''''''''''''''''''''''''''
                        Unload Me
                        
                    End If
            End If
        
        End If
    End With
Else
        For i = 1 To 1000
             DoEvents
        Next i
        SwipeProcess = True
End If
Exit Sub
SWipeError:
 ' frmBadSwipe.Show 1
'ImBadSwipe.Visible = True
Call LogClick("Swipe Error", Err.Description)
'If Err.Number = 7 Then
 'TERMINATE APPLICATION
  End
'End If
'MsgBox Err.Description
'TimerBadSwipe.Interval = 4000
'TimerBadSwipe.Enabled = True
'SwipeProcess = True

End Sub


Private Sub USBHID1_CardDataError()
SwipeProcess = False
'lresult = USBHID1.USBSwipe_Command("02", output)
'USBHID1.PortOpen = True
Call LogClick("Swipe", "Bad Swipe")

'frmBadSwipe.Show 1
ImBadSwipe.Visible = True
For i = 1 To 10
     DoEvents
Next i

TimerBadSwipe.interval = 3000
TimerBadSwipe.Enabled = True
SwipeProcess = True

End Sub


