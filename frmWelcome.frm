VERSION 5.00
Object = "{158336E7-3FF3-456E-912C-5985E9BBED24}#1.2#0"; "MTUSBHIDSwipe.ocx"
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   12120
   ClientLeft      =   2235
   ClientTop       =   855
   ClientWidth     =   14160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   12120
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ctlUSBHID.USBHID USBHID1 
      Left            =   135
      Top             =   3420
      _ExtentX        =   1826
      _ExtentY        =   1296
   End
   Begin VB.Timer TimerClick 
      Left            =   165
      Top             =   2835
   End
   Begin VB.Timer Timer3 
      Left            =   180
      Top             =   2205
   End
   Begin VB.Timer Timer2 
      Left            =   195
      Top             =   1695
   End
   Begin VB.Image imMagStrip 
      Height          =   1695
      Left            =   735
      Picture         =   "frmWelcome.frx":17A7
      Top             =   6030
      Width           =   2160
   End
   Begin VB.Label lOffLine 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   225
      TabIndex        =   7
      Top             =   8010
      Width           =   240
   End
   Begin VB.Label TXTbought 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   75
      TabIndex        =   6
      Top             =   8340
      Width           =   375
   End
   Begin VB.Image click2image 
      Height          =   3000
      Left            =   11085
      Picture         =   "frmWelcome.frx":2D24
      Top             =   8790
      Width           =   3000
   End
   Begin VB.Image Click1Image 
      Appearance      =   0  'Flat
      Height          =   3000
      Left            =   30
      Picture         =   "frmWelcome.frx":3FD1
      Top             =   8880
      Width           =   3000
   End
   Begin VB.Label Disclaimer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "It is the policy of the NAIAS to conduct random security checks with Metal Detectors and of bags and packages."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   30
      TabIndex        =   5
      Top             =   10200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Discover 
      Height          =   585
      Left            =   8340
      Picture         =   "frmWelcome.frx":527E
      Top             =   8415
      Width           =   1845
   End
   Begin VB.Image ImAmex 
      Height          =   795
      Left            =   7350
      Picture         =   "frmWelcome.frx":7194
      Top             =   8355
      Width           =   1065
   End
   Begin VB.Image ImMC 
      Height          =   825
      Left            =   4410
      Picture         =   "frmWelcome.frx":7665
      Top             =   8370
      Width           =   1410
   End
   Begin VB.Image ImaVisa 
      Height          =   840
      Left            =   5925
      Picture         =   "frmWelcome.frx":8216
      Top             =   8355
      Width           =   1395
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   480
      Top             =   8055
      Width           =   1545
   End
   Begin VB.Image ImgStop 
      Height          =   1560
      Left            =   0
      Top             =   -15
      Width           =   3090
   End
   Begin VB.Label LabelClose 
      BackStyle       =   0  'Transparent
      Height          =   810
      Left            =   300
      TabIndex        =   4
      Top             =   1770
      Width           =   765
   End
   Begin VB.Image imSwipe 
      Height          =   2265
      Left            =   10815
      Picture         =   "frmWelcome.frx":8DF1
      Top             =   5475
      Width           =   3240
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Swipe card to begin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1020
      Index           =   3
      Left            =   -330
      TabIndex        =   3
      Top             =   6495
      Width           =   14715
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004694C1&
      Height          =   1020
      Index           =   2
      Left            =   1530
      TabIndex        =   2
      Top             =   4470
      Width           =   11715
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004694C1&
      Height          =   525
      Index           =   1
      Left            =   2835
      TabIndex        =   1
      Top             =   3855
      Width           =   8175
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004694C1&
      Height          =   1020
      Index           =   0
      Left            =   2835
      TabIndex        =   0
      Top             =   2805
      Width           =   8175
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imNo As Integer
Dim numcycles As Integer

Dim ClickClick1 As Boolean
Dim ClickClick2 As Boolean




Private Sub Click1Image_Click()
ClickClick1 = True
TimerClick.Interval = 5000
TimerClick.Enabled = True
Timer2.Interval = 0
Timer2.Enabled = False


End Sub

Private Sub click2image_Click()
If ClickClick1 Then
        TimerClick.Interval = 0
        TimerClick.Enabled = False
        ClickClick1 = False
        Timer2.Interval = 0
        Timer2.Enabled = False
        fTicketmation.Show 1
        'Unload Me
        
End If
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim rsTickets As Recordset
Dim i As Integer
Dim rsTicketStock As Recordset
Dim OpenForms As Integer
Dim rsBatch As Recordset

ReDim TicketTypes(5, 0)
ReDim TotalTicketSold(0)
On Error GoTo appError
Timer2.Enabled = False
bStartup = False
If bTicketmation Then
    bDEMO = False
    bTicketmation = False
End If
Set rsBatch = db.OpenRecordset("Select * from BATCH where BatchUpload=No")

If rsBatch.EOF Then
    lOffLine.Caption = ""
Else
    lOffLine.Caption = "*"
End If
'If bDebug Then MsgBox "ready to unload frmPINPAD"
''frmPINPAD.Dsintx1.ReleaseConn
'For i = 1 To 1000
'    OpenForms = DoEvents   ' Yield to operating system.
'Next i
'If bDebug Then MsgBox "OpenForms:" & OpenForms
'Unload frmPINPAD
'
'If bDebug Then MsgBox "frmPINPAD unloaded"
'
'Load frmPINPAD
'If bDebug Then MsgBox "frmPINPAD Loaded"

'Clear SALE

'Dsintx1.SetTfield "hostname", Text1(1).Text
'Dsintx1.SetTfield "hostport", Text1(2).Text
'Dsintx1.SetTfield "DoSSL", "Y"

'Dim xx As String
'Dim transResult As Boolean
'
'Dsintx1.SetTfield "hostname", "o13d.eigendev.com"
'Dsintx1.SetTfield "hostport", "2000"
'Dsintx1.SetTfield "DoSSL", "Y"
'Dsintx1.SetTfield "PinPadPort", "4"
'
'Dsintx1.InitTransaction
'
'Dsintx1.SetTfield "Trans_Code", "01"
'Dsintx1.SetTfield "Track2_Acc", "M4501123456789=03030?"
'Dsintx1.SetTfield "Amount_1", "263"
'Dsintx1.SetTfield "Eigen_ID", "Eigen_ID"
'Dsintx1.SetTfield "Term_ID", "Term_ID"
'
'transResult = Dsintx1.ProcessTransaction
ClickClick1 = False
ClickClick2 = False

If VisaMCOnly Then
    ImAmex.Visible = False
End If
If AllowDiscover Then
    Discover.Visible = True
Else
    Discover.Visible = False
End If
numcycles = 0
Set rsTicketStock = db.OpenRecordset("Select * from TICKET_STOCK")
If rsTicketStock(0) <= 10 Then
    lblWelcome(3).Caption = "Sorry - Out of ticket stock.  Please use other terminal"
    lblWelcome(3).FontSize = 18
    lblWelcome(2).Caption = ""
    lblWelcome(1).Caption = ""
    lblWelcome(0).Caption = ""
    imSwipe.Visible = False
    imMagStrip.Visible = False
    Disclaimer.Visible = False
    Label1.Visible = False
Else
    TXTbought.Caption = rsTicketStock(0)
    TicketStock = rsTicketStock(0)
    bChange = False
    If Not bServer Then
        lblWelcome(3).Caption = "Sorry - this machine is out of service.  Please use other terminal"
        lblWelcome(3).FontSize = 16
        lblWelcome(2).Caption = ""
        lblWelcome(1).Caption = ""
        lblWelcome(0).Caption = ""
        imSwipe.Visible = False
        imMagStrip.Visible = False
        Label1.Visible = False
    Else
        Set rs = db.OpenRecordset("Select ShowName,EventReferenceName,EventCity from SHOWS where showID =" & ShowID)
        lblWelcome(2).Caption = rs("ShowName")
        ShowReference = rs("EventReferenceName")
        EventCity = rs("EventCity")
'        If EventCity = "DETROIT" Then
'            Disclaimer.Visible = True
'        Else
            Disclaimer.Visible = False
'        End If
'        Timer2.Interval = 1000
'        Timer2.Enabled = True
        USBHID1.PortOpen = True
        If Not (USBHID1.PortOpen) Then
         MsgBox "Couldn't open HID Swipe Reader"
         End
        End If

    End If
End If
Exit Sub
appError:
    lblWelcome(3).Caption = "Sorry - this machine is out of service.  Please use other terminal"
    lblWelcome(3).FontSize = 16
     Label1.Visible = False

    imSwipe.Visible = False
    imMagStrip.Visible = False

End Sub


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim ShiftDown, AltDown, CtrlDown, Txt
    If bDebug Or bDEMO Then
        AltDown = (Shift And vbAltMask) > 0
        If AltDown Then
            SwipTime = Now()
            'AccountNumber = ";6011300111103=16071300000000100100001?"
            Track1 = "%B3728 944600 04019^GOODMAN/HS               ^0805040682662?"
            AccountNumber = ";4007000000027=16071300000000100100001?"
            CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            CardExpire = Right(CardExpire, 2) & "/" & Left(CardExpire, 2)
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            CardName = "Goodman/HS"
            If bDebug Then MsgBox "Card Type:" & Left(CardNumber, 1) & "," & VisaMCOnly
            If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
                'Card is a debit
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            'ElseIf VisaMCOnly And Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4") And (CardType <> "120" And CardType <> "220") Then
            ElseIf VisaMCOnly And Left(CardNumber, 1) = "3" Then
                'Card is an AMEX
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                If NotAllowDebit Then
                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA AND MASTERCARD."
                Else
                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD OR DEBIT CARDS."
                End If
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 4) = "6011" And Not AllowDiscover Then
                'Card is a debit
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DISCOVER CARDS.  PLEASE USE A DIFFERENT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4" And Left(CardNumber, 1) <> "3" And Left(CardNumber, 4) <> "6011" And CardType <> "120" And CardType <> "220" Then
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "PLEASE USE A VALID CARD."
    '                If AllowDiscover And Not VisaMCOnly And Not NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD, AMEX OR DEBIT CARDS."
    '                ElseIf AllowDiscover And Not VisaMCOnly And NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD OR AMEX."
    '                ElseIf AllowDiscover And VisaMCOnly And Not NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD OR DEBIT CARDS."
    '                ElseIf AllowDiscover And VisaMCOnly And Not NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD OR DEBIT CARDS."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            End If
            'CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            'Track1 = .Dsintx1.GetTfield("Track1")
            'If Len(Track1) > 1 Then
            '    CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
            'Else
            '    CardName = ""
            'End If
            '.Dsintx1.InitTransaction
           ' .Dsintx1.ReleaseConn
    '    'MsgBox CardNumber
    '    'MsgBox CardExpire
            db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
            Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
            SessionID = rs(0)
            Call LogClick("Welcome", "SwipeCard")
            If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
            'Me.Hide
            frmSpecifyTickets.Show
            Unload Me
        Else
            numcycles = numcycles + 1
            If numcycles > 1 And bMultiple Then
                frmWelcomeMult.Show
                'Unload frmPINPAD
                Unload Me
            Else
                Timer2.Interval = 100
                Timer2.Enabled = True
            End If
            End If
    
''' NOT SURE           Unload Me
        'End If
    End If
End Sub


Private Sub Image3_Click()

End Sub

Private Sub ImgStop_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim ShiftDown, AltDown, CtrlDown, Txt
    AltDown = (Shift And vbAltMask) > 0
    If AltDown Then
            frmMainAN.Show
            'frmMenu.Show
            Unload Me
    End If

End Sub


Private Sub imSwipe_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim ShiftDown, AltDown, CtrlDown, Txt
    If bDebug Or bDEMO Then
        AltDown = (Shift And vbAltMask) > 0
        If AltDown Then
            SwipTime = Now()
            'AccountNumber = ";6011300111103=16071300000000100100001?"
            Track1 = "%B3728 944600 04019^GOODMAN/HS               ^0805040682662?"
            AccountNumber = ";4007000000027=16071300000000100100001?"
            CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            CardExpire = Right(CardExpire, 2) & "/" & Left(CardExpire, 2)
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            CardName = "Goodman/HS"
            If bDebug Then MsgBox "Card Type:" & Left(CardNumber, 1) & "," & VisaMCOnly
            If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
                'Card is a debit
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            'ElseIf VisaMCOnly And Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4") And (CardType <> "120" And CardType <> "220") Then
            ElseIf VisaMCOnly And Left(CardNumber, 1) = "3" Then
                'Card is an AMEX
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                If NotAllowDebit Then
                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA AND MASTERCARD."
                Else
                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD OR DEBIT CARDS."
                End If
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 4) = "6011" And Not AllowDiscover Then
                'Card is a debit
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DISCOVER CARDS.  PLEASE USE A DIFFERENT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4" And Left(CardNumber, 1) <> "3" And Left(CardNumber, 4) <> "6011" And CardType <> "120" And CardType <> "220" Then
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "PLEASE USE A VALID CARD."
    '                If AllowDiscover And Not VisaMCOnly And Not NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD, AMEX OR DEBIT CARDS."
    '                ElseIf AllowDiscover And Not VisaMCOnly And NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD OR AMEX."
    '                ElseIf AllowDiscover And VisaMCOnly And Not NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD OR DEBIT CARDS."
    '                ElseIf AllowDiscover And VisaMCOnly And Not NotAllowDebit Then
    '                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD, DISCOVERY CARD OR DEBIT CARDS."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            End If
            'CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            'Track1 = .Dsintx1.GetTfield("Track1")
            'If Len(Track1) > 1 Then
            '    CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
            'Else
            '    CardName = ""
            'End If
            '.Dsintx1.InitTransaction
           ' .Dsintx1.ReleaseConn
    '    'MsgBox CardNumber
    '    'MsgBox CardExpire
            db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
            Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
            SessionID = rs(0)
            Call LogClick("Welcome", "SwipeCard")
            If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
            'Me.Hide
            frmSpecifyTickets.Show
            Unload Me
        Else
            numcycles = numcycles + 1
            If numcycles > 1 And bMultiple Then
                frmWelcomeMult.Show
                'Unload frmPINPAD
                Unload Me
            Else
                Timer2.Interval = 100
                Timer2.Enabled = True
            End If
            End If
    
''' NOT SURE           Unload Me
        'End If
    End If

End Sub


Private Sub LabelClose_DblClick()
End

End Sub



Private Sub Timer2_Timer()
Dim xreturn As Boolean
'On Error GoTo procError
xreturn = False
Timer2.Interval = 0
Timer2.Enabled = False

''With frmPINPAD
''    .Dsintx1.SetTfield "PinPadPort", 1
''    '.Dsintx1.SetTfield "PinPadPort", 5
''    .Dsintx1.SetTfield "NoWriteLog", "Y"
''    .Dsintx1.InitTransaction
''    .Dsintx1.SetTfield "ReadCard", "y"
''    xreturn = .Dsintx1.ProcessTransaction
''    'If bDebug Then MsgBox "xreturn:" & xreturn
''    If xreturn Then
''        SwipTime = Now()
''
'        If AllowDiscover Then
'            AccountNumber = ";6011300111103=16071300000000100100001?"
'            AllowDiscover = False
'        Else
'            AccountNumber = .Dsintx1.GetTfield("Track2_Acc")
 '       End If
''        If Len(AccountNumber) > 0 Then
''            CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
''            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
''
''            If bDebug Then MsgBox "Card Type:" & Left(CardNumber, 1) & "," & VisaMCOnly
''            If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
''                'Card is a debit
''                lblWelcome(3).FontSize = 16
''                lblWelcome(3).ForeColor = &HFF&
''                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
''                Timer3.Interval = 3000
''                Timer3.Enabled = True
''                Exit Sub
''            'ElseIf VisaMCOnly And Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4") And (CardType <> "120" And CardType <> "220") Then
''            ElseIf VisaMCOnly And Left(CardNumber, 1) = "3" Then
''                'Card is an AMEX
''                lblWelcome(3).FontSize = 16
''                lblWelcome(3).ForeColor = &HFF&
''                If NotAllowDebit Then
''                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA AND MASTERCARD."
''                Else
''                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD OR DEBIT CARDS."
''                End If
''                Timer3.Interval = 3000
''                Timer3.Enabled = True
''                Exit Sub
''            ElseIf Left(CardNumber, 4) = "6011" And Not AllowDiscover Then
''                'Card is a debit
''                lblWelcome(3).FontSize = 16
''                lblWelcome(3).ForeColor = &HFF&
''                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DISCOVER CARDS.  PLEASE USE A DIFFERENT CARD."
''                Timer3.Interval = 3000
''                Timer3.Enabled = True
''                Exit Sub
''            ElseIf Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4" And Left(CardNumber, 1) <> "3" And Left(CardNumber, 4) <> "6011" And CardType <> "120" And CardType <> "220" Then
''                lblWelcome(3).FontSize = 16
''                lblWelcome(3).ForeColor = &HFF&
''                lblWelcome(3).Caption = "PLEASE USE A VALID CARD."
''                Timer3.Interval = 3000
''                Timer3.Enabled = True
''                Exit Sub
''            End If
''            'CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
''            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
''            Track1 = .Dsintx1.GetTfield("Track1")
''            If bDebug Then MsgBox "Track1=" & Track1
''            If bDebug Then MsgBox "Track2=" & AccountNumber
''            If Len(Track1) > 1 Then
''                CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
''            Else
''                CardName = ""
''            End If
''            .Dsintx1.InitTransaction
''            .Dsintx1.ReleaseConn
''    '    'MsgBox CardNumber
''    '    'MsgBox CardExpire
''            db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
''            Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
''            SessionID = rs(0)
''            Call LogClick("Welcome", "SwipeCard")
''            If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
''            'Me.Hide
''            frmSpecifyTickets.Show
''            Unload Me
''        Else
''            numcycles = numcycles + 1
''            If numcycles > 1 And bMultiple Then
''                frmWelcomeMult.Show
''                'Unload frmPINPAD
''                Unload Me
''            Else
''                Timer2.Interval = 100
''                Timer2.Enabled = True
''            End If
''        End If
''    End If
''    Exit Sub
''End With
''procError:
''    If bDebug Then MsgBox "Process Error:" & Err.Description
''    numcycles = numcycles + 1
''    If numcycles > 0 And bMultiple Then
''        frmWelcomeMult.Show
''        'Unload frmPINPAD
''        Unload Me
''    Else
''        Timer2.Interval = 100
''        Timer2.Enabled = True
''    End If
End Sub


Private Sub Timer3_Timer()
lblWelcome(3).ForeColor = &H8000&
lblWelcome(3).FontSize = 28
lblWelcome(3).Caption = "Swipe Credit Card to Begin"
Timer3.Interval = 0
Timer3.Enabled = False

Timer2.Interval = 100
Timer2.Enabled = True

End Sub


Private Sub TimerClick_Timer()
ClickClick1 = False
TimerClick.Interval = 0
Timer2.Interval = 1000
Timer2.Enabled = True

End Sub


Private Sub USBHID1_CardDataChanged()
Dim abTrack1() As Byte
Dim abTrack2() As Byte
Dim abTrack3() As Byte
Dim i As Integer
Dim strOutput As String
''Text1.Text = ""
With USBHID1
If (.GetTrack(1) <> "" Or .GetTrack(2) <> "" Or .GetTrack(3) <> "") Then
        SwipTime = Now()
''
        AccountNumber = .GetTrack(2)
        If Len(AccountNumber) > 0 Then
            CardType = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 5, 3)
            CardNumber = Mid(AccountNumber, 2, InStr(1, AccountNumber, "=") - 2)
''
            If bDebug Then MsgBox "Card Type:" & Left(CardNumber, 1) & "," & VisaMCOnly
            If (CardType = "120" Or CardType = "220") And NotAllowDebit Then
                'Card is a debit
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DEBIT CARDS.  PLEASE USE A CREDIT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            'ElseIf VisaMCOnly And Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4") And (CardType <> "120" And CardType <> "220") Then
            ElseIf VisaMCOnly And Left(CardNumber, 1) = "3" Then
                'Card is an AMEX
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                If NotAllowDebit Then
                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA AND MASTERCARD."
                Else
                    lblWelcome(3).Caption = "THIS MACHINE ONLY ACCEPTS VISA, MASTERCARD OR DEBIT CARDS."
                End If
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 4) = "6011" And Not AllowDiscover Then
                'Card is a debit
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "THIS MACHINE DOES NOT ACCEPT DISCOVER CARDS.  PLEASE USE A DIFFERENT CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            ElseIf Left(CardNumber, 1) <> "5" And Left(CardNumber, 1) <> "4" And Left(CardNumber, 1) <> "3" And Left(CardNumber, 4) <> "6011" And CardType <> "120" And CardType <> "220" Then
                lblWelcome(3).FontSize = 16
                lblWelcome(3).ForeColor = &HFF&
                lblWelcome(3).Caption = "PLEASE USE A VALID CARD."
                Timer3.Interval = 3000
                Timer3.Enabled = True
                Exit Sub
            End If
            CardExpire = Mid(AccountNumber, InStr(1, AccountNumber, "=") + 1, 4)
            Track1 = .GetTrack(1)
            If bDebug Then MsgBox "Track1=" & Track1
            If bDebug Then MsgBox "Track2=" & AccountNumber
            If Len(Track1) > 1 Then
                CardName = Trim(Mid(Track1, InStr(1, Track1, "^") + 1, InStr((InStr(1, Track1, "^") + 1), Track1, "^") - InStr(1, Track1, "^") - 1))
            Else
                CardName = .GetLName & "/" & .GetFName
            End If
''            .Dsintx1.InitTransaction
''            .Dsintx1.ReleaseConn
            USBHID1.PortOpen = False

            db.Execute "Insert into SESSION (SessionStart) values(#" & Now() & "#)"
            Set rs = db.OpenRecordset("Select max(SessionID) from SESSION")
            SessionID = rs(0)
            Call LogClick("Welcome", "SwipeCard")
            If bDebug Then MsgBox "Card Info:" & CardNumber & "," & CardExpire
            'Me.Hide
            frmSpecifyTickets.Show
            Unload Me
''        Else
''            numcycles = numcycles + 1
''            If numcycles > 1 And bMultiple Then
''                frmWelcomeMult.Show
''                'Unload frmPINPAD
''                Unload Me
''            Else
''                Timer2.Interval = 100
''                Timer2.Enabled = True
''            End If
        End If
''    End If


''With USBHID1
''If (.GetTrack(1) <> "" Or .GetTrack(2) <> "" Or .GetTrack(3) <> "") Then
''
''    Text1.Text = Text1.Text & "CardData = " & .CardData
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "CardDataMasked = " & .CardDataMasked
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "FirstName = " & .GetFName
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "LastName = " & .GetLName
''    Text1.Text = Text1.Text & vbCrLf
''
''    Text1.Text = Text1.Text & "Track1Data = " & .GetTrack(1)
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track2Data = " & .GetTrack(2)
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track3Data = " & .GetTrack(3)
''    Text1.Text = Text1.Text & vbCrLf
''
''    Text1.Text = Text1.Text & "Track1DataLength = " & .Track1Length
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track2DataLength = " & .Track2Length
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track3DataLength = " & .Track3Length
''    Text1.Text = Text1.Text & vbCrLf
''
''    Text1.Text = Text1.Text & "FirstNameMasked = " & .GetFNameMasked
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "LastNameMasked = " & .GetLNameMasked
''    Text1.Text = Text1.Text & vbCrLf
''
''    Text1.Text = Text1.Text & "Track1DataLengthmasked = " & .Track1LengthMasked
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track2DataLengthmasked = " & .Track2LengthMasked
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track3DataLengthmasked = " & .Track3LengthMasked
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track1DataMasked = " & .GetTrackMasked(1)
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track2DataMasked = " & .GetTrackMasked(2)
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Track3DataMasked = " & .GetTrackMasked(3)
''    Text1.Text = Text1.Text & vbCrLf
''
''    Text1.Text = Text1.Text & "DUKPTKSN = " & .DUKPTKSN
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "SessionID = " & .SessionID
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "SessionIDHex = " & .SessionIDHexString
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "MagnePrint = " & .MPrintData
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "MagnePrintHex = " & .MagnePrintDataHexString
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "MagnePrint Length = " & .MPrintLen
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "Sequence Number = " & .SequenceNumber
''    Text1.Text = Text1.Text & vbCrLf
''    Text1.Text = Text1.Text & "ReaderID = " & .ReaderID
''    Text1.Text = Text1.Text & vbCrLf
End If
End With

End Sub


Private Sub USBHID1_CardDataError()
MsgBox ("Error Reading Card. Please try again")

End Sub


